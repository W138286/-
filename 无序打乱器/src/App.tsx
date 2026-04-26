/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useCallback, useRef, useEffect, useMemo } from 'react';
import { 
  Download, 
  Shuffle, 
  Trash2,
  FileText,
  Settings2,
  Sparkles,
  Sun,
  Moon,
  Type
} from 'lucide-react';
import { 
  Document, 
  Packer, 
  Table, 
  TableRow, 
  TableCell, 
  Paragraph, 
  TextRun, 
  WidthType,
  AlignmentType,
  VerticalAlign,
  BorderStyle,
} from 'docx';
import { saveAs } from 'file-saver';
import { cn } from './lib/utils';
import {
  DndContext,
  closestCenter,
  KeyboardSensor,
  PointerSensor,
  useSensor,
  useSensors,
  DragEndEvent
} from '@dnd-kit/core';
import {
  arrayMove,
  SortableContext,
  sortableKeyboardCoordinates,
  verticalListSortingStrategy,
  useSortable
} from '@dnd-kit/sortable';
import { CSS } from '@dnd-kit/utilities';
import { GripVertical, X, Layout } from 'lucide-react';

// --- Components ---
function SortableItem({ id, children, className }: { id: string, children: React.ReactNode, className?: string, key?: any }) {
  const {
    attributes,
    listeners,
    setNodeRef,
    transform,
    transition,
    isDragging
  } = useSortable({ id });

  const style = {
    transform: CSS.Transform.toString(transform),
    transition,
    zIndex: isDragging ? 10 : 1,
    opacity: isDragging ? 0.5 : 1,
  };

  return (
    <div ref={setNodeRef} style={style} className={cn("relative group", className)}>
      <div {...attributes} {...listeners} className="absolute left-0 top-1/2 -translate-y-1/2 -translate-x-full pr-2 opacity-0 group-hover:opacity-100 cursor-grab active:cursor-grabbing transition-opacity">
        <GripVertical size={14} className="text-zinc-400" />
      </div>
      {children}
    </div>
  );
}

// --- Helpers ---
const POS_LIST = ['n', 'v', 'adj', 'adv', 'pron', 'prep', 'conj', 'art', 'num', 'int', 'aux', 'mod'];

const getTodayDate = () => {
  const now = new Date();
  const year = now.getFullYear();
  const month = String(now.getMonth() + 1).padStart(2, '0');
  const day = String(now.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
};

const shuffleArray = <T,>(array: T[]): T[] => {
  const newArr = [...array];
  for (let i = newArr.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [newArr[i], newArr[j]] = [newArr[j], newArr[i]];
  }
  return newArr;
};

// --- Words Input Area ---

function TextAreaInput({ 
  value, 
  onChange, 
  onScroll, 
  placeholder, 
  className,
  themeStyles,
  textareaRef,
  onKeyDown
}: {
  value: string;
  onChange: (e: React.ChangeEvent<HTMLTextAreaElement>) => void;
  onScroll?: (e: React.UIEvent<HTMLTextAreaElement>) => void;
  onKeyDown?: (e: React.KeyboardEvent<HTMLTextAreaElement>) => void;
  placeholder?: string;
  className?: string;
  themeStyles: any;
  textareaRef?: React.RefObject<HTMLTextAreaElement>;
}) {
  return (
    <div 
      className={cn("flex-1 flex cursor-text h-full", className)}
      onClick={() => textareaRef?.current?.focus()}
    >
      <textarea
        ref={textareaRef}
        value={value}
        onChange={onChange}
        onScroll={onScroll}
        onKeyDown={onKeyDown}
        placeholder={placeholder}
        spellCheck={false}
        wrap="off"
        className={cn(
          "block m-0 border-0 flex-1 px-6 py-6 text-sm focus:outline-none resize-none font-sans transition-colors bg-transparent min-h-full whitespace-nowrap",
          themeStyles.inputText
        )}
        style={{ lineHeight: '26px' }}
      />
    </div>
  );
}

// --- Line Navigation Helpers ---
const getLineInfo = (text: string, pos: number) => {
  const linesBefore = text.slice(0, pos).split('\n');
  const lineIndex = linesBefore.length - 1;
  const colIndex = linesBefore[lineIndex].length;
  const allLines = text.split('\n');
  return {
    lineIndex,
    colIndex,
    isAtStart: colIndex === 0,
    isAtEnd: colIndex === (allLines[lineIndex]?.length || 0),
    totalLines: allLines.length
  };
};

const getPosAtLineEdge = (text: string, lineIndex: number, atEnd: boolean) => {
  const lines = text.split('\n');
  if (lineIndex < 0) return 0;
  if (lineIndex >= lines.length) return text.length;
  let pos = 0;
  for (let i = 0; i < lineIndex; i++) {
    pos += lines[i].length + 1;
  }
  return atEnd ? pos + (lines[lineIndex]?.length || 0) : pos;
};

class ErrorBoundary extends React.Component<{children: React.ReactNode}, {hasError: boolean, error: Error | null}> {
  constructor(props: {children: React.ReactNode}) {
    super(props);
    this.state = { hasError: false, error: null };
  }

  static getDerivedStateFromError(error: Error) {
    return { hasError: true, error };
  }

  componentDidCatch(error: Error, errorInfo: React.ErrorInfo) {
    console.error("ErrorBoundary caught an error", error, errorInfo);
  }

  render() {
    if (this.state.hasError) {
      return (
        <div style={{ padding: '40px', fontFamily: 'sans-serif', color: '#ff4444', backgroundColor: '#fff', minHeight: '100vh' }}>
          <h2>应用渲染崩溃 / Application Crashed</h2>
          <p>很抱歉，应用在渲染时遇到了未知错误。请截图此页面反馈：</p>
          <pre style={{ background: '#f5f5f5', padding: '20px', whiteSpace: 'pre-wrap', color: '#333', fontSize: '12px', border: '1px solid #ddd', borderRadius: '8px' }}>
            {this.state.error && this.state.error.stack ? this.state.error.stack : this.state.error?.toString()}
          </pre>
          <button 
            onClick={() => window.location.reload()}
            style={{ marginTop: '20px', padding: '10px 20px', border: 'none', background: '#3b82f6', color: 'white', borderRadius: '4px', cursor: 'pointer' }}
          >
            尝试重新加载 (Reload)
          </button>
        </div>
      );
    }
    return this.props.children;
  }
}

const storage = {
  getItem(key: string): string | null {
    try {
      return window.localStorage.getItem(key);
    } catch (e) {
      return null;
    }
  },
  setItem(key: string, value: string): void {
    try {
      window.localStorage.setItem(key, value);
    } catch (e) {}
  },
  removeItem(key: string): void {
    try {
      window.localStorage.removeItem(key);
    } catch (e) {}
  }
};

export default function AppWrapper() {
  return (
    <ErrorBoundary>
      <App />
    </ErrorBoundary>
  );
}

function App() {
  const [input, setInput] = useState<string>('');
  const [posInput, setPosInput] = useState<string>('');
  const [transformationInput, setTransformationInput] = useState<string>('');
  const [sentenceInput, setSentenceInput] = useState<string>('');
  const [exportMode, setExportMode] = useState<'4-col' | '6-col' | 'review'>(() => {
    if (typeof window !== 'undefined') {
      const saved = storage.getItem('app-export-mode');
      return (saved === '4-col' || saved === '6-col' || saved === 'review') ? saved as any : '4-col';
    }
    return '4-col';
  });
  const [isExamMode, setIsExamMode] = useState<boolean>(() => {
    if (typeof window !== 'undefined') {
      return storage.getItem('app-exam-mode') === 'true';
    }
    return false;
  });
  const [isExporting, setIsExporting] = useState(false);
  const [showPreview, setShowPreview] = useState(false);
  const [isDragging, setIsDragging] = useState(false);
  const [previewItems, setPreviewItems] = useState<{ id: string; term: string; pos: string; answer: string }[]>([]);
  const [previewTransformations, setPreviewTransformations] = useState<{ id: string; content: string }[]>([]);
  const [previewSentences, setPreviewSentences] = useState<{ id: string; content: string }[]>([]);
  const [columnWidths, setColumnWidths] = useState<{ word: number; pos: number; reviewRatio: number }>({ word: 50, pos: 50, reviewRatio: 0.7 });
  const [isSavingDefaults, setIsSavingDefaults] = useState(false);
  const [theme, setTheme] = useState<'light' | 'dark' | 'sepia'>(() => {
    if (typeof window !== 'undefined') {
      const saved = storage.getItem('app-theme');
      return (saved === 'dark' || saved === 'light' || saved === 'sepia') ? saved as any : 'light';
    }
    return 'light';
  });
  const [answerInput, setAnswerInput] = useState<string>('');
  const [isSyncDelete, setIsSyncDelete] = useState<boolean>(() => {
    if (typeof window !== 'undefined') {
      return storage.getItem('app-sync-delete') !== 'false';
    }
    return true;
  });
  const [autoSaveInterval, setAutoSaveInterval] = useState<number>(() => {
    if (typeof window !== 'undefined') {
      const saved = storage.getItem('app-autosave-interval');
      return saved ? parseInt(saved) : 5; // Default 5 mins
    }
    return 5;
  });
  const [enterJump, setEnterJump] = useState<boolean>(() => {
    if (typeof window !== 'undefined') {
      const saved = storage.getItem('app-enter-jump');
      return saved !== null ? saved === 'true' : true;
    }
    return true;
  });

  const [sepiaIntensity, setSepiaIntensity] = useState<number>(() => {
    if (typeof window !== 'undefined') {
      const saved = storage.getItem('app-sepia-intensity');
      return saved ? Number(saved) : 50;
    }
    return 50;
  });

  // Memoized line arrays to prevent splitting on every render for large files
  const inputLineArray = useMemo(() => Array.from({ length: input.split('\n').length }), [input]);
  const posLineArray = useMemo(() => Array.from({ length: posInput.split('\n').length }), [posInput]);
  const answerLineArray = useMemo(() => Array.from({ length: answerInput.split('\n').length }), [answerInput]);

  useEffect(() => {
    storage.setItem('app-sepia-intensity', String(sepiaIntensity));
  }, [sepiaIntensity]);

  const themeStyles = useMemo(() => {
    const intensity = sepiaIntensity / 100;
    
    // Sepia calculation: interpolate between light sepia and deep sepia
    // Base: #F4ECD8 (244, 236, 216) -> Deep: #D2C2A4 (210, 194, 164)
    const r = Math.round(244 - (244 - 210) * intensity);
    const g = Math.round(236 - (236 - 194) * intensity);
    const b = Math.round(216 - (216 - 164) * intensity);
    const sepiaBgColor = `rgb(${r}, ${g}, ${b})`;

    return {
      light: {
        bg: "bg-[#F5F5F7] text-[#1D1D1F]",
        border: "border-black",
        borderMuted: "border-black/10",
        subBg: "bg-gray-100",
        innerBg: "bg-white/50",
        inputBg: "bg-white",
        inputText: "text-black",
        inputBorder: "border-black",
        mutedText: "text-zinc-500",
        accent: "bg-black text-white",
        toggleBg: "bg-gray-200",
        lineNumColor: "text-black/40"
      },
      dark: {
        bg: "bg-[#0A0A0B] text-[#E4E4E7]",
        border: "border-zinc-800",
        borderMuted: "border-zinc-800",
        subBg: "bg-zinc-950",
        innerBg: "bg-zinc-950/40",
        inputBg: "bg-zinc-900",
        inputText: "text-zinc-300",
        inputBorder: "border-zinc-700",
        mutedText: "text-zinc-500",
        accent: "bg-white text-black",
        toggleBg: "bg-zinc-800",
        lineNumColor: "text-zinc-700"
      },
      sepia: {
        bg: "text-[#5B4636]",
        customBg: sepiaBgColor,
        border: "border-[#D2C2A4]",
        borderMuted: "border-[#D2C2A4]/50",
        subBg: "bg-[#EBE0C5]",
        innerBg: "bg-[#EBE0C5]/40",
        inputBg: "bg-transparent",
        inputText: "text-[#5B4636]",
        inputBorder: "border-[#D2C2A4]",
        mutedText: "text-[#8C7B60]",
        accent: "bg-[#5B4636] text-[#F4ECD8]",
        toggleBg: "bg-[#E1D4B9]",
        lineNumColor: "text-[#5B4636]/40"
      }
    }[theme];
  }, [theme, sepiaIntensity]);

  const parseExtractedText = (text: string) => {
    const lines = text.split('\n').map(line => line.trim()).filter(line => line.length > 0);
    const parsedWords: string[] = [];
    const parsedPos: string[] = [];
    const parsedAnswers: string[] = [];

    lines.forEach(line => {
      const posRegex = new RegExp(`(?:^|\\s)(${POS_LIST.join('|')}|phr|phrase)\\.?(?=\\s|[^a-zA-Z]|$)`, 'i');
      const match = line.match(posRegex);
      
      if (match) {
        const posIndex = match.index!;
        let posToken = match[1].toLowerCase();
        if (posToken === 'phr' || posToken === 'phrase') posToken = 'phr';
        if (!posToken.endsWith('.')) posToken += '.';
        
        const wordPart = line.substring(0, posIndex).trim();
        const ansPart = line.substring(posIndex + match[0].length).trim();
        
        parsedWords.push(wordPart);
        parsedPos.push(posToken);
        parsedAnswers.push(ansPart);
      } else {
        const chineseMatch = line.match(/[^\x00-\x7F]/);
        if (chineseMatch) {
          const index = chineseMatch.index!;
          parsedWords.push(line.substring(0, index).trim());
          parsedPos.push("");
          parsedAnswers.push(line.substring(index).trim());
        } else {
          parsedWords.push(line);
          parsedPos.push("");
          parsedAnswers.push("");
        }
      }
    });

    if (parsedWords.some(w => w)) setInput(parsedWords.join('\n'));
    if (parsedPos.some(p => p)) setPosInput(parsedPos.join('\n'));
    if (parsedAnswers.some(a => a)) setAnswerInput(parsedAnswers.join('\n'));
  };

  const processFile = async (file?: File | null) => {
    if (!file) return;

    try {
      let text = '';
      if (file.name.endsWith('.docx')) {
        const mammoth = await import('mammoth');
        const arrayBuffer = await file.arrayBuffer();
        const result = await mammoth.extractRawText({ arrayBuffer });
        text = result.value;
      } else {
        text = await file.text();
      }
      parseExtractedText(text);
    } catch (err) {
      console.error("File read error:", err);
      alert("文件读取失败，请确保格式正确。");
    }
  };

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    processFile(e.target.files?.[0]);
  };

  const onDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(true);
  };

  const onDragLeave = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
  };

  const onDrop = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
    const file = e.dataTransfer.files?.[0];
    if (file) {
      processFile(file);
    }
  };

  const syncDeletions = (oldVal: string, newVal: string, targetSetter: (val: string | ((v: string) => string)) => void, targetVal: string) => {
    if (!isSyncDelete) return;
    
    // Simple line-based detection
    const oldLines = oldVal.split('\n');
    const newLines = newVal.split('\n');
    
    if (newLines.length < oldLines.length) {
      // Find deleted line index
      let deletedIdx = -1;
      for (let i = 0; i < oldLines.length; i++) {
        if (oldLines[i] !== newLines[i]) {
          // Verify if it's actually removed (shifted)
          if (oldLines[i+1] === newLines[i]) {
            deletedIdx = i;
            break;
          }
        }
      }
      
      if (deletedIdx !== -1) {
        const targetLines = targetVal.split('\n');
        if (targetLines.length > deletedIdx) {
          targetLines.splice(deletedIdx, 1);
          targetSetter(targetLines.join('\n'));
        }
      }
    }
  };
  const [exportSettings, setExportSettings] = useState(() => {
    const defaultSettings = {
      fontSize: 10.5,
      font: 'Times New Roman',
      lineSpacing: 1.5,
      columnSpacing: 1.3,
      isCentered: true,
      maxCharsPerLine: 25,
      headerChoice: '中文' as '英文' | '中文'
    };
    if (typeof window !== 'undefined') {
      const saved = storage.getItem('app-export-settings');
      if (saved) {
        try {
          return { ...defaultSettings, ...JSON.parse(saved) };
        } catch (e) {
          console.error("Failed to parse saved settings", e);
        }
      }
    }
    return defaultSettings;
  });
  const scrollRef = useRef<HTMLDivElement>(null);
  const wordInputRef = useRef<HTMLTextAreaElement>(null);
  const posInputRef = useRef<HTMLTextAreaElement>(null);
  const ansInputRef = useRef<HTMLTextAreaElement>(null);
  const previewTableRef = useRef<HTMLDivElement>(null);

  const sensors = useSensors(
    useSensor(PointerSensor, {
      activationConstraint: {
        distance: 8,
      },
    }),
    useSensor(KeyboardSensor, {
      coordinateGetter: sortableKeyboardCoordinates,
    })
  );

  const handleDragEndItems = (event: DragEndEvent) => {
    const { active, over } = event;
    if (over && active.id !== over.id) {
      setPreviewItems((items) => {
        const oldIndex = items.findIndex((i) => i.id === active.id);
        const newIndex = items.findIndex((i) => i.id === over.id);
        return arrayMove(items, oldIndex, newIndex);
      });
    }
  };

  const handleDragEndTrans = (event: DragEndEvent) => {
    const { active, over } = event;
    if (over && active.id !== over.id) {
      setPreviewTransformations((items) => {
        const oldIndex = items.findIndex((i) => i.id === active.id);
        const newIndex = items.findIndex((i) => i.id === over.id);
        return arrayMove(items, oldIndex, newIndex);
      });
    }
  };

  const handleDragEndSent = (event: DragEndEvent) => {
    const { active, over } = event;
    if (over && active.id !== over.id) {
      setPreviewSentences((items) => {
        const oldIndex = items.findIndex((i) => i.id === active.id);
        const newIndex = items.findIndex((i) => i.id === over.id);
        return arrayMove(items, oldIndex, newIndex);
      });
    }
  };

  const preparePreview = () => {
    const rawItems = processInputToFlatList(input).filter(i => i.length > 0);
    const rawPos = processInputToFlatList(posInput);
    const rawAns = processInputToFlatList(answerInput);
    const trans = processInputToFlatList(transformationInput).filter(i => i.length > 0);
    const sents = processInputToFlatList(sentenceInput).filter(i => i.length > 0);

    setPreviewItems(shuffleArray(rawItems.map((term, i) => ({ 
      id: `row-${i}-${Date.now()}`, 
      term, 
      pos: rawPos[i] || '',
      answer: rawAns[i] || ''
    }))));
    setPreviewTransformations(shuffleArray(trans.map((c, i) => ({ id: `trans-${i}-${Date.now()}`, content: c }))));
    setPreviewSentences(shuffleArray(sents.map((c, i) => ({ id: `sent-${i}-${Date.now()}`, content: c }))));
    
    // Set default widths based on mode
    if (exportMode === '4-col') {
      const colS = exportSettings.columnSpacing;
      const wPos = Math.max(5, Math.min(45, 25 * (colS / (1 + colS) * 2)));
      setColumnWidths(prev => ({ ...prev, word: 50 - wPos, pos: wPos }));
    } else {
      setColumnWidths(prev => ({ ...prev, word: 16.6, pos: 16.6 })); // Placeholder logic for 6-col simplified
    }
    
    setShowPreview(true);
  };

  const handleDividerDrag = (e: React.MouseEvent, type: 'main' | 'review' = 'main') => {
    if (!previewTableRef.current) return;
    
    const container = previewTableRef.current;
    const startX = e.clientX;
    const startWordWidth = columnWidths.word;
    const startReviewRatio = columnWidths.reviewRatio;
    const containerWidth = container.offsetWidth;

    const onMouseMove = (moveEvent: MouseEvent) => {
      const deltaX = moveEvent.clientX - startX;
      const deltaPercent = (deltaX / containerWidth) * 100;
      
      if (type === 'review') {
        // Adjust the split between term and pos within the 'word' width
        // newRatio * wordWidth = startRatio * wordWidth + deltaPercent
        let newRatio = startReviewRatio + (deltaPercent / startWordWidth);
        newRatio = Math.max(0.1, Math.min(0.9, newRatio));
        setColumnWidths(prev => ({ ...prev, reviewRatio: newRatio }));
      } else {
        const unitWidth = exportMode === '6-col' ? 33.333 : 50;
        const scalingFactor = 50 / unitWidth; // Normalize to 50% unit scale
        
        let newWordWidth = Math.max(5, Math.min(45, startWordWidth + (deltaPercent * scalingFactor)));
        setColumnWidths(prev => ({ ...prev, word: newWordWidth, pos: 50 - newWordWidth }));
      }
    };

    const onMouseUp = () => {
      window.removeEventListener('mousemove', onMouseMove);
      window.removeEventListener('mouseup', onMouseUp);
      document.body.style.cursor = '';
    };

    window.addEventListener('mousemove', onMouseMove);
    window.addEventListener('mouseup', onMouseUp);
    document.body.style.cursor = 'col-resize';
  };

  // Persistence for theme
  useEffect(() => {
    storage.setItem('app-theme', theme);
    document.documentElement.setAttribute('data-theme', theme);
  }, [theme]);

  // Auto-save logic
  useEffect(() => {
    if (autoSaveInterval <= 0) return;

    const interval = setInterval(() => {
      const data = {
        input,
        posInput,
        answerInput,
        transformationInput,
        sentenceInput,
        exportSettings,
        exportMode,
        isExamMode,
        timestamp: Date.now()
      };
      storage.setItem('app-autosave-data', JSON.stringify(data));
      storage.setItem('app-last-autosave-time', String(Date.now()));
      console.log('Document auto-saved');
    }, autoSaveInterval * 1000);

    return () => clearInterval(interval);
  }, [autoSaveInterval, input, posInput, answerInput, transformationInput, sentenceInput, exportSettings, exportMode, isExamMode]);

  // Monthly cleanup
  useEffect(() => {
    const lastCleanup = storage.getItem('app-last-cleanup');
    const now = Date.now();
    const oneMonth = 30 * 24 * 60 * 60 * 1000;

    if (!lastCleanup || now - parseInt(lastCleanup) > oneMonth) {
      // Cleanup old autosaves but keep the latest one if it's very fresh
      storage.removeItem('app-autosave-data');
      storage.setItem('app-last-cleanup', String(now));
    }

    // Load last autosave if user wants or if it exists?
    // For now, we rely on standard state, but we could offer restoration.
  }, []);

  // Confirmation before exit
  useEffect(() => {
    const handleBeforeUnload = (e: BeforeUnloadEvent) => {
      if (input.trim() || posInput.trim()) {
        e.preventDefault();
        e.returnValue = '';
      }
    };
    window.addEventListener('beforeunload', handleBeforeUnload);
    return () => window.removeEventListener('beforeunload', handleBeforeUnload);
  }, [input, posInput]);

  const toggleTheme = () => {
    setTheme(prev => {
      if (prev === 'light') return 'dark';
      if (prev === 'dark') return 'sepia';
      return 'light';
    });
  };

  const saveDefaultSettings = () => {
    storage.setItem('app-export-settings', JSON.stringify(exportSettings));
    // storage.setItem('app-export-mode', exportMode); // Excluded based on user request
    storage.setItem('app-exam-mode', String(isExamMode));
    storage.setItem('app-enter-jump', String(enterJump));
    storage.setItem('app-sync-delete', String(isSyncDelete));
    storage.setItem('app-autosave-interval', String(autoSaveInterval));
    storage.setItem('app-sepia-intensity', String(sepiaIntensity));
    
    // Quick visual feedback
    setIsSavingDefaults(true);
    setTimeout(() => {
      setIsSavingDefaults(false);
    }, 2000);
  };

  const wrapText = (text: string, max: number) => {
    if (max <= 0) return text;
    if (text.length <= max) return text;
    
    const words = text.split(' ');
    const lines: string[] = [];
    let currentLine = "";

    for (const word of words) {
      if ((currentLine + word).length <= max) {
        currentLine += (currentLine === "" ? "" : " ") + word;
      } else {
        if (currentLine !== "") lines.push(currentLine);
        // If the single word is longer than max, we must split it
        if (word.length > max) {
          let remaining = word;
          while (remaining.length > max) {
            lines.push(remaining.slice(0, max));
            remaining = remaining.slice(max);
          }
          currentLine = remaining;
        } else {
          currentLine = word;
        }
      }
    }
    if (currentLine !== "") lines.push(currentLine);
    return lines.join('\n');
  };

  const processInputToFlatList = useCallback((raw: string): string[] => {
    return raw.split('\n').map(line => line.trim());
  }, []);

  const handleExport = async (itemsData?: typeof previewItems, transData?: typeof previewTransformations, sentData?: typeof previewSentences) => {
    const finalItems = itemsData || shuffleArray(processInputToFlatList(input).filter(i => i.length > 0).map((term, i) => ({ 
      id: `row-${i}`, 
      term, 
      pos: processInputToFlatList(posInput)[i] || '',
      answer: processInputToFlatList(answerInput)[i] || ''
    })));
    const finalTrans = transData || shuffleArray(processInputToFlatList(transformationInput).filter(i => i.length > 0).map((c, i) => ({ id: `trans-${i}`, content: c })));
    const finalSentences = sentData || shuffleArray(processInputToFlatList(sentenceInput).filter(i => i.length > 0).map((c, i) => ({ id: `sent-${i}`, content: c })));
    
    if (finalItems.length === 0 && finalTrans.length === 0 && finalSentences.length === 0) return;
    setIsExporting(true);

    try {
      const getPara = (text: string, isHeader = false, isCentered?: boolean) => {
        const processedText = wrapText(text, exportSettings.maxCharsPerLine);
        const alignment = isCentered ?? exportSettings.isCentered;
        return new Paragraph({
          alignment: alignment ? AlignmentType.CENTER : AlignmentType.LEFT,
          children: processedText.split('\n').map((line, idx, array) => (
            new TextRun({
              text: line,
              size: exportSettings.fontSize * 2,
              font: isHeader ? 'Microsoft YaHei' : exportSettings.font,
              bold: isHeader,
              break: idx < array.length - 1 ? 1 : undefined
            })
          )),
          spacing: { line: Math.round(exportSettings.lineSpacing * 240) }
        });
      };

      const createCell = (text: string, width: number, isHeader = false, isCentered?: boolean) => {
        const trimmed = text.trim();
        const isSlash = ['/', '\\', '／'].includes(trimmed);
        return new TableCell({
          children: [getPara(isSlash ? '' : text, isHeader, isCentered)],
          verticalAlign: VerticalAlign.CENTER,
          width: { size: width, type: WidthType.PERCENTAGE },
          borders: {
            top: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
            bottom: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
            left: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
            right: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
            ...(isSlash ? { diagonalDown: { style: BorderStyle.SINGLE, size: 2, color: "000000" } } : {})
          }
        });
      };

      const sections: any[] = [];
      const tableRows: TableRow[] = [];
      const colS = exportSettings.columnSpacing;
      const choice = exportSettings.headerChoice;

      if (finalItems.length > 0) {
        if (exportMode === '4-col') {
          tableRows.push(new TableRow({
            children: [
              createCell('测试提示', 25, true), createCell(choice, 25, true),
              createCell('测试提示', 25, true), createCell(choice, 25, true),
            ]
          }));
          const W_SPACE = itemsData ? columnWidths.pos : Math.max(5, Math.min(45, 25 * (colS / (1 + colS) * 2)));
          const W_ITEM = itemsData ? columnWidths.word : 50 - W_SPACE;
          for (let i = 0; i < finalItems.length; i += 2) {
            tableRows.push(new TableRow({
              children: [
                createCell(finalItems[i]?.term || '', W_ITEM), createCell(finalItems[i]?.answer || '', W_SPACE),
                createCell(finalItems[i+1]?.term || '', W_ITEM), createCell(finalItems[i+1]?.answer || '', W_SPACE),
              ]
            }));
          }
        } else if (exportMode === '6-col') {
          // New 6-col logic: Word + POS, 3 items per row
          tableRows.push(new TableRow({
            children: [
              createCell('测试提示', 16.6, true), createCell('词性', 16.6, true),
              createCell('测试提示', 16.6, true), createCell('词性', 16.6, true),
              createCell('测试提示', 16.6, true), createCell('词性', 16.6, true),
            ]
          }));
          const W_POS = 16.6 * (colS / 2);
          const W_ITEM = 33.3 - W_POS;
          for (let i = 0; i < finalItems.length; i += 3) {
            tableRows.push(new TableRow({
              children: [
                createCell(finalItems[i]?.term || '', W_ITEM), createCell(finalItems[i]?.pos || '', W_POS),
                createCell(finalItems[i+1]?.term || '', W_ITEM), createCell(finalItems[i+1]?.pos || '', W_POS),
                createCell(finalItems[i+2]?.term || '', W_ITEM), createCell(finalItems[i+2]?.pos || '', W_POS),
              ]
            }));
          }
        } else {
          // 'review' mode: Word + POS + Answer, 2 items per row
          tableRows.push(new TableRow({
            children: [
              createCell('测试提示', 16.6, true), createCell('词性', 16.6, true), createCell(choice, 16.6, true),
              createCell('测试提示', 16.6, true), createCell('词性', 16.6, true), createCell(choice, 16.6, true),
            ]
          }));
          const W_POS = 16.6 * (colS / 2);
          const W_CHIN = 16.6 * 2 - W_POS;
          const W_ITEM = 16.6;
          for (let i = 0; i < finalItems.length; i += 2) {
            tableRows.push(new TableRow({
              children: [
                createCell(finalItems[i]?.term || '', W_ITEM), createCell(finalItems[i]?.pos || '', W_POS), createCell(finalItems[i]?.answer || '', W_CHIN),
                createCell(finalItems[i+1]?.term || '', W_ITEM), createCell(finalItems[i+1]?.pos || '', W_POS), createCell(finalItems[i+1]?.answer || '', W_CHIN),
              ]
            }));
          }
        }
        sections.push(new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows: tableRows }));
      }

      if (isExamMode && finalTrans.length > 0) {
        sections.push(new Paragraph({ children: [new TextRun({ text: '\n变形', bold: true, size: 24, font: 'Microsoft YaHei' })] }));
        const transTableRows: TableRow[] = [];
        for (let i = 0; i < finalTrans.length; i += 2) {
          const left = finalTrans[i]?.content, right = finalTrans[i+1]?.content;
          transTableRows.push(new TableRow({
            children: [
              createCell(left ? left + "____________________" : "", 50, false, false),
              createCell(right ? right + "____________________" : "", 50, false, false),
            ]
          }));
        }
        sections.push(new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows: transTableRows }));
      }

      if (isExamMode && finalSentences.length > 0) {
        sections.push(new Paragraph({ children: [new TextRun({ text: '\n句子考查', bold: true, size: 24, font: 'Microsoft YaHei' })] }));
        finalSentences.forEach((s) => {
          const cleanText = s.content.replace(/^\d+[\.\s、]+/, '').trim();
          sections.push(new Paragraph({
            children: [new TextRun({ text: cleanText, size: exportSettings.fontSize * 2, font: exportSettings.font })],
            spacing: { before: 200 }
          }));
          sections.push(new Paragraph({
            children: [new TextRun({ text: "____________________________________________________________________________________", size: exportSettings.fontSize * 2 })],
            spacing: { after: 200 }
          }));
        });
      }

      const doc = new Document({ sections: [{ children: sections }] });
      const blob = await Packer.toBlob(doc);
      saveAs(blob, `${getTodayDate()}.docx`);
    } catch (error) {
      console.error('Export failed:', error);
    } finally {
      setIsExporting(false);
    }
  };

  const handlePOSChange = (e: React.ChangeEvent<HTMLTextAreaElement>) => {
    let value = e.target.value;
    const start = e.target.selectionStart;
    const end = e.target.selectionEnd;
    
    // Standardize slashes
    value = value.replace(/[／\\|]/g, '/');
    
    // Trigger on Space or Newline to finalize the word
    if (value.length > posInput.length) {
      const lastChar = value[value.length - 1];
      if (lastChar === ' ' || lastChar === '\n') {
        const lines = value.split('\n');
        // If it was a newline, the word we care about is on the line BEFORE the last one
        const targetLineIndex = lastChar === '\n' ? lines.length - 2 : lines.length - 1;
        if (targetLineIndex >= 0) {
          const targetLine = lines[targetLineIndex];
          const match = targetLine.match(/([a-zA-Z]+)$/);
          if (match) {
            const word = match[1].toLowerCase();
            if (POS_LIST.includes(word)) {
              const beforePos = value.lastIndexOf(match[1]);
              const base = value.slice(0, beforePos);
              const suffix = lastChar === ' ' ? '. ' : '.\n';
              const newVal = base + word + suffix;
              setPosInput(newVal);
              
              // Set selection after state update
              setTimeout(() => {
                if (posInputRef.current) {
                   const newOffset = newVal.length;
                   posInputRef.current.setSelectionRange(newOffset, newOffset);
                }
              }, 0);
              return;
            }
          }
        }
      }
    }
    setPosInput(value);
  };

  const handleScroll = (e: React.UIEvent<HTMLTextAreaElement>, ref: React.RefObject<HTMLDivElement>) => {
    if (ref.current) ref.current.scrollTop = e.currentTarget.scrollTop;
  };

  return (
    <div className={cn(
      "min-h-screen font-sans p-6 md:p-12 transition-all duration-300",
      themeStyles.bg
    )} style={theme === 'sepia' ? { backgroundColor: themeStyles.customBg } : {}}>
      <div className="max-w-6xl mx-auto">
        <header className={cn(
          "flex items-end justify-between mb-16 border-b pb-8 transition-colors",
          themeStyles.border
        )}>
          <div>
            <h1 className="text-4xl font-light tracking-tight mb-2 uppercase">无序打乱器</h1>
            <p className={cn("text-sm font-medium", theme === 'light' ? "text-zinc-800" : theme === 'sepia' ? "text-[#5B4636]/80" : "text-zinc-500")}>
              极简单词无序测试制作工具。 Produced by Werther with Gemini 3.1.
            </p>
          </div>
          <div className="flex items-center gap-6">
            <div className={cn("flex p-1 rounded-2xl border", theme === 'light' ? "bg-zinc-200/50 border-zinc-300" : theme === 'sepia' ? "bg-[#E1D4B9] border-[#D2C2A4]" : "bg-zinc-800/50 border-zinc-700")}>
               <button 
                onClick={() => setTheme('light')}
                className={cn("p-2 rounded-xl transition-all", theme === 'light' ? "bg-white text-black shadow-sm" : "text-zinc-500 hover:text-black")}
               >
                 <Sun size={14} />
               </button>
               <button 
                onClick={() => setTheme('sepia')}
                className={cn("p-2 rounded-xl transition-all", theme === 'sepia' ? "bg-[#5B4636] text-[#F4ECD8] shadow-sm" : "text-zinc-500 hover:text-[#5B4636]")}
               >
                 <Sparkles size={14} />
               </button>
               <button 
                onClick={() => setTheme('dark')}
                className={cn("p-2 rounded-xl transition-all", theme === 'dark' ? "bg-zinc-800 text-white shadow-sm border border-white/10" : "text-zinc-500 hover:text-white")}
               >
                 <Moon size={14} />
               </button>
            </div>
            <div className="text-right hidden md:block">
              <span className={cn("text-[10px] uppercase tracking-widest font-bold", theme === 'light' ? "text-gray-500" : theme === 'sepia' ? "text-[#8C7B60]" : "text-zinc-700")}>
                v5.0 PRO Edition
              </span>
            </div>
          </div>
        </header>

        <main className="grid grid-cols-1 lg:grid-cols-[1fr_450px] gap-12 items-start">
          <section className="space-y-6">
            <div className="flex items-center justify-between">
              <div className={cn("flex items-center gap-2 text-xs font-bold uppercase tracking-widest", themeStyles.inputText)}>
                <FileText size={14} /> 录入区域
              </div>
              <div className="flex gap-4">
                  <label 
                    className={cn("relative text-[10px] uppercase font-bold transition-all px-3 py-1.5 rounded-xl border active:scale-95 cursor-pointer flex items-center gap-2 shadow-sm", isDragging ? "border-dashed border-cyan-500 bg-cyan-50/10 text-cyan-500 scale-105" : (theme === 'light' ? "text-zinc-800 border-black bg-white hover:bg-gray-50" : theme === 'sepia' ? "text-[#5B4636] border-[#D2C2A4] bg-[#F4ECD8] hover:bg-[#EBE0C5]" : "text-zinc-300 border-zinc-700 bg-zinc-900 hover:bg-zinc-800 hover:text-white"))}
                    onDragOver={onDragOver}
                    onDragLeave={onDragLeave}
                    onDrop={onDrop}
                  >
                    <Download size={12} className="rotate-180" /> 上传文档
                    <input type="file" accept=".txt,.docx" className="hidden" onChange={handleFileUpload} />
                  </label>
                  <button 
                  onClick={() => {
                    setInput("joint efforts\nbe stuck at home\nfacilitate\nInteraction\nreclaim\ncut down\nlessen");
                    setPosInput("n.\nphrase\nv.\nn.\nv.\nv./phr.\nv.");
                    if (isExamMode) {
                      setTransformationInput("comprehend 变形形容词\nenhance 变形名词");
                      setSentenceInput("1. 我写信是为了推荐.... 我的理由如下。\n2. 正如谚语所说：“赠人玫瑰，手有余香。”");
                    }
                  }}
                  className={cn("text-[10px] uppercase font-bold transition-all px-3 py-1.5 rounded-xl border active:scale-95 shadow-sm", theme === 'light' ? "text-zinc-800 border-black bg-white hover:bg-gray-50" : theme === 'sepia' ? "text-[#5B4636] border-[#D2C2A4] bg-[#F4ECD8] hover:bg-[#EBE0C5]" : "text-zinc-300 border-zinc-700 bg-zinc-900 hover:bg-zinc-800 hover:text-white")}
                >
                  载入示例
                </button>
                <button 
                  onClick={() => { setInput(''); setPosInput(''); }} 
                  className={cn("transition-all p-1.5 rounded-xl border active:scale-95", theme === 'light' ? "border-black text-zinc-400 hover:text-red-500" : theme === 'sepia' ? "border-[#D2C2A4] text-[#8C7B60] hover:text-red-500" : "border-zinc-800 text-gray-300 hover:text-red-500")}
                >
                  <Trash2 size={16} />
                </button>
              </div>
            </div>

            <div className={cn(
              "relative flex rounded-3xl border overflow-hidden min-h-[700px] transition-all",
              themeStyles.inputBg, themeStyles.border, "shadow-xl shadow-black/5"
            )} style={theme === 'sepia' ? { backgroundColor: themeStyles.customBg } : {}}>
              {exportMode === '4-col' ? (
                <div className="flex flex-col flex-1 relative group">
                  <div className={cn("px-6 py-3 border-b text-[10px] font-bold uppercase tracking-widest", themeStyles.subBg, themeStyles.border, themeStyles.inputText)}>测试提示</div>
                  <div className="flex-1 flex min-h-0">
                    <div ref={scrollRef} className={cn("w-12 border-r py-6 text-right pr-3 select-none overflow-hidden transition-colors pointer-events-none block", themeStyles.subBg, themeStyles.border)}>
                      {inputLineArray.map((_, i) => (
                        <div key={i} style={{ height: '26px', lineHeight: '26px' }} className={cn("text-[11px] flex items-center justify-end font-sans shrink-0", themeStyles.lineNumColor)}>{i + 1}</div>
                      ))}
                    </div>
                    <TextAreaInput
                      value={input}
                      onChange={(e) => setInput(e.target.value)}
                      onScroll={(e) => { if (scrollRef.current) scrollRef.current.scrollTop = e.currentTarget.scrollTop; }}
                      placeholder="输入清单，每行一个..."
                      themeStyles={themeStyles}
                      className=""
                    />
                  </div>
                </div>
              ) : (
                <div className={cn(
                  "flex-1 grid divide-x", 
                  exportMode === '6-col' ? "grid-cols-[1fr_1fr]" : "grid-cols-[1fr_1fr_1fr]",
                  themeStyles.border
                )}>
                  <div className="flex flex-col relative group">
                    <div className={cn("px-6 py-3 border-b text-[10px] font-bold uppercase tracking-widest", themeStyles.subBg, themeStyles.border, themeStyles.inputText)}>测试提示</div>
                    <div className="flex-1 flex min-h-0">
                      <div ref={scrollRef} className={cn("w-10 border-r py-6 text-right pr-2 select-none overflow-hidden transition-colors pointer-events-none block", themeStyles.subBg, themeStyles.border)}>
                        {inputLineArray.map((_, i) => (
                          <div key={i} style={{ height: '26px', lineHeight: '26px' }} className={cn("text-[10px] flex items-center justify-end font-sans shrink-0", themeStyles.lineNumColor)}>{i + 1}</div>
                        ))}
                      </div>
                      <TextAreaInput
                        textareaRef={wordInputRef}
                        value={input}
                        onChange={(e) => {
                          const oldVal = input;
                          const newVal = e.target.value;
                          setInput(newVal);
                          syncDeletions(oldVal, newVal, setPosInput, posInput);
                          syncDeletions(oldVal, newVal, setAnswerInput, answerInput);
                        }}
                        onScroll={(e) => {
                          const top = e.currentTarget.scrollTop;
                          if (scrollRef.current) scrollRef.current.scrollTop = top;
                          if (posInputRef.current && posInputRef.current.scrollTop !== top) posInputRef.current.scrollTop = top;
                          if (ansInputRef.current && ansInputRef.current.scrollTop !== top) ansInputRef.current.scrollTop = top;
                        }}
                        placeholder="单词..."
                        themeStyles={themeStyles}
                        className=""
                        onKeyDown={(e) => {
                          if (enterJump && !e.shiftKey && e.key === 'Enter') {
                            e.preventDefault();
                            const { lineIndex } = getLineInfo(e.currentTarget.value, e.currentTarget.selectionStart);
                            if (posInputRef.current) {
                              // Ensure target has that line
                              const targetLines = posInput.split('\n');
                              while (targetLines.length <= lineIndex) targetLines.push('');
                              if (targetLines.length > posInput.split('\n').length) setPosInput(targetLines.join('\n'));
                              
                              setTimeout(() => {
                                if (posInputRef.current) {
                                  posInputRef.current.focus();
                                  const targetPos = getPosAtLineEdge(posInputRef.current.value, lineIndex, false);
                                  posInputRef.current.setSelectionRange(targetPos, targetPos);
                                }
                              }, 0);
                            }
                          }
                        }}
                      />
                    </div>
                  </div>
                    <div className="flex flex-col relative group">
                      <div className={cn("px-6 py-3 border-b text-[10px] font-bold uppercase tracking-widest", themeStyles.subBg, themeStyles.border, themeStyles.inputText)}>
                        词性
                      </div>
                      <div className="flex-1 flex min-h-0">
                        <TextAreaInput
                          textareaRef={posInputRef}
                      value={posInput}
                      onChange={(e) => {
                        handlePOSChange(e);
                        const oldVal = posInput;
                        const newVal = e.target.value;
                        syncDeletions(oldVal, newVal, setInput, input);
                        syncDeletions(oldVal, newVal, setAnswerInput, answerInput);
                      }}
                      onScroll={(e) => {
                        const top = e.currentTarget.scrollTop;
                        if (scrollRef.current) scrollRef.current.scrollTop = top;
                        if (wordInputRef.current && wordInputRef.current.scrollTop !== top) wordInputRef.current.scrollTop = top;
                        if (ansInputRef.current && ansInputRef.current.scrollTop !== top) ansInputRef.current.scrollTop = top;
                      }}
                      className=""
                      placeholder="词性..."
                      themeStyles={themeStyles}
                      onKeyDown={(e) => {
                        if (enterJump && !e.shiftKey && e.key === 'Enter') {
                          e.preventDefault();
                          const { lineIndex } = getLineInfo(e.currentTarget.value, e.currentTarget.selectionStart);
                          
                          // 1. Manually add dot if needed before jumping
                          const lines = posInput.split('\n');
                          const curLine = lines[lineIndex] || '';
                          const match = curLine.match(/([a-zA-Z]+)$/);
                          if (match && POS_LIST.includes(match[1].toLowerCase())) {
                            lines[lineIndex] = curLine + '.';
                            setPosInput(lines.join('\n'));
                          }

                          // 2. Jump logic
                          if (exportMode === 'review' && ansInputRef.current) {
                            const targetLines = answerInput.split('\n');
                            while (targetLines.length <= lineIndex) targetLines.push('');
                            if (targetLines.length > answerInput.split('\n').length) setAnswerInput(targetLines.join('\n'));
                            
                            setTimeout(() => {
                              if (ansInputRef.current) {
                                ansInputRef.current.focus();
                                const targetPos = getPosAtLineEdge(ansInputRef.current.value, lineIndex, false);
                                ansInputRef.current.setSelectionRange(targetPos, targetPos);
                              }
                            }, 0);
                          } else if (wordInputRef.current) {
                            const nextLine = lineIndex + 1;
                            const targetLines = input.split('\n');
                            if (targetLines.length <= nextLine) {
                              setInput(input + '\n');
                            }
                            
                            setTimeout(() => {
                              if (wordInputRef.current) {
                                wordInputRef.current.focus();
                                const targetPos = getPosAtLineEdge(wordInputRef.current.value, nextLine, false);
                                wordInputRef.current.setSelectionRange(targetPos, targetPos);
                              }
                            }, 0);
                          }
                        }
                      }}
                    />
                    </div>
                  </div>
                  {exportMode === 'review' && (
                    <div className="flex flex-col relative group">
                      <div className={cn("px-6 py-3 border-b text-[10px] font-bold uppercase tracking-widest", themeStyles.subBg, themeStyles.border, themeStyles.inputText)}>
                        答案
                      </div>
                      <div className="flex-1 flex min-h-0">
                        <TextAreaInput
                          textareaRef={ansInputRef}
                        value={answerInput}
                        onChange={(e) => {
                          const oldVal = answerInput;
                          const newVal = e.target.value;
                          setAnswerInput(newVal);
                          syncDeletions(oldVal, newVal, setInput, input);
                          syncDeletions(oldVal, newVal, setPosInput, posInput);
                        }}
                        onScroll={(e) => {
                          const top = e.currentTarget.scrollTop;
                          if (scrollRef.current) scrollRef.current.scrollTop = top;
                          if (wordInputRef.current && wordInputRef.current.scrollTop !== top) wordInputRef.current.scrollTop = top;
                          if (posInputRef.current && posInputRef.current.scrollTop !== top) posInputRef.current.scrollTop = top;
                        }}
                        className=""
                        placeholder="中文答案..."
                        themeStyles={themeStyles}
                        onKeyDown={(e) => {
                          if (enterJump && !e.shiftKey && e.key === 'Enter') {
                            e.preventDefault();
                            const { lineIndex } = getLineInfo(e.currentTarget.value, e.currentTarget.selectionStart);
                            const nextLine = lineIndex + 1;
                            
                            if (wordInputRef.current) {
                              const targetLines = input.split('\n');
                              if (targetLines.length <= nextLine) {
                                setInput(input + '\n');
                              }
                              
                              setTimeout(() => {
                                if (wordInputRef.current) {
                                  wordInputRef.current.focus();
                                  const targetPos = getPosAtLineEdge(wordInputRef.current.value, nextLine, false);
                                  wordInputRef.current.setSelectionRange(targetPos, targetPos);
                                }
                              }, 0);
                            }
                          }
                        }}
                      />
                      </div>
                    </div>
                  )}
                </div>
              )}
            </div>

            {isExamMode && (
              <div className="space-y-6 animate-in slide-in-from-bottom-4 duration-500">
                <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                  <div className="space-y-3">
                    <div className={cn("flex items-center gap-2 text-xs font-bold uppercase tracking-widest", themeStyles.inputText)}>
                      <Type size={14} /> 变形考查
                    </div>
                    <div className={cn("rounded-3xl border p-4", themeStyles.inputBg, themeStyles.border)}>
                       <textarea
                        value={transformationInput}
                        onChange={(e) => setTransformationInput(e.target.value)}
                        placeholder="输入变形单词，如：&#10;comprehend 变形形容词&#10;enhance 变形名词"
                        className={cn("w-full min-h-[150px] bg-transparent resize-none text-sm font-sans focus:outline-none leading-relaxed transition-colors", themeStyles.inputText)}
                      />
                    </div>
                  </div>
                  <div className="space-y-3">
                    <div className={cn("flex items-center gap-2 text-xs font-bold uppercase tracking-widest", themeStyles.inputText)}>
                      <FileText size={14} /> 句子考查 (自动编号)
                    </div>
                    <div className={cn("rounded-3xl border p-4", themeStyles.inputBg, themeStyles.border)}>
                       <textarea
                        value={sentenceInput}
                        onChange={(e) => {
                          let val = e.target.value;
                          // If it's the very first character and not a number, prepend '1. '
                          if (val.length === 1 && /^[^\d]/.test(val)) {
                            val = '1. ' + val;
                          }
                          setSentenceInput(val);
                        }}
                        onKeyDown={(e) => {
                          if (e.key === 'Enter' && !e.shiftKey) {
                            const textarea = e.currentTarget;
                            const { lineIndex } = getLineInfo(textarea.value, textarea.selectionStart);
                            const lines = sentenceInput.split('\n');
                            const currentLine = lines[lineIndex];
                            
                            // Detect numbering: "1.", "1 ", "1、"
                            const match = currentLine.match(/^(\d+)[\.\s、]/);
                            if (match) {
                              e.preventDefault();
                              const num = parseInt(match[1]);
                              const nextNum = num + 1;
                              const insertText = `\n${nextNum}. `;
                              const start = textarea.selectionStart;
                              const before = sentenceInput.slice(0, start);
                              const after = sentenceInput.slice(textarea.selectionEnd);
                              
                              setSentenceInput(before + insertText + after);
                              
                              // Use requestAnimationFrame or setTimeout to set cursor position after re-render
                              setTimeout(() => {
                                const newPos = start + insertText.length;
                                textarea.focus();
                                textarea.setSelectionRange(newPos, newPos);
                              }, 0);
                            }
                          }
                        }}
                        placeholder="输入考查句子，回车自动换行编号..."
                        className={cn("w-full min-h-[150px] bg-transparent resize-none text-sm font-mono focus:outline-none leading-relaxed", theme !== 'dark' ? "text-black" : "text-zinc-300")}
                      />
                    </div>
                  </div>
                </div>
              </div>
            )}
          </section>

          <aside className="space-y-8 pt-10">
            <div className={cn("rounded-[32px] p-8 border space-y-8 shadow-sm transition-all", themeStyles.inputBg, themeStyles.border)}>
              
              {/* Primary Actions - Top Frequency */}
              <div className="grid grid-cols-1 gap-4">
                <div className={cn("flex flex-col p-1 rounded-2xl border", themeStyles.subBg, themeStyles.border)}>
                  <button
                    onClick={preparePreview}
                    disabled={!input.trim()}
                    className={cn(
                      "w-full px-5 py-6 rounded-xl text-xs font-bold transition-all flex items-center justify-center border shadow-xl active:scale-[0.98]", 
                      !input.trim()
                        ? (theme === 'light' ? "bg-gray-50/50 text-zinc-400 border-black/10 cursor-not-allowed" : theme === 'sepia' ? "bg-[#EBE0C5]/50 text-[#8C7B60]/50 border-[#D2C2A4]/20 cursor-not-allowed" : "bg-zinc-900/50 text-zinc-700 border-zinc-800/50 cursor-not-allowed")
                        : theme === 'light'
                        ? "bg-black text-white border-black hover:bg-zinc-800"
                        : theme === 'sepia'
                        ? "bg-[#5B4636] text-[#F4ECD8] border-[#5B4636] hover:bg-[#4A392C]"
                        : "bg-zinc-900 text-zinc-300 border-zinc-700 hover:bg-zinc-800 hover:text-white"
                    )}
                  >
                    <Layout size={16} className="inline mr-2" /> 预览文稿
                  </button>
                </div>
                <div className={cn("flex flex-col p-1 rounded-2xl border", themeStyles.subBg, themeStyles.border)}>
                  <button
                    onClick={() => {
                      const rawItems = processInputToFlatList(input).filter(i => i.length > 0);
                      const rawPos = processInputToFlatList(posInput);
                      const rawAns = processInputToFlatList(answerInput);
                      
                      const items: {term: string, pos: string, answer: string}[] = rawItems.map((term, i) => ({ 
                        term, 
                        pos: rawPos[i] || '', 
                        answer: rawAns[i] || '' 
                      }));
                      
                      const shuffled = shuffleArray(items);
                      
                      setInput(shuffled.map(p => p.term).join('\n'));
                      setPosInput(shuffled.map(p => p.pos).join('\n'));
                      setAnswerInput(shuffled.map(p => p.answer).join('\n'));
                    }}
                    disabled={!input.trim()}
                    className={cn(
                      "w-full px-5 py-4 rounded-xl text-[11px] font-bold transition-all flex items-center justify-center border shadow-sm", 
                      theme === 'light' 
                        ? (input.trim() ? "bg-white text-black border-black hover:bg-black hover:text-white" : "bg-gray-50/50 text-zinc-600 border-black/20 cursor-not-allowed")
                        : theme === 'sepia'
                        ? (input.trim() ? "bg-[#F4ECD8] text-[#5B4636] border-[#5B4636] hover:bg-[#5B4636] hover:text-[#F4ECD8]" : "bg-[#EBE0C5]/30 text-[#8C7B60] border-[#D2C2A4]/20 cursor-not-allowed")
                        : (input.trim() ? "bg-zinc-900 text-zinc-300 border-zinc-700 hover:bg-zinc-800 hover:text-white" : "bg-zinc-900/50 text-zinc-700 border-zinc-800/50 cursor-not-allowed")
                    )}
                  >
                    <Shuffle size={14} className={cn("inline mr-2")} /> 随机打乱
                  </button>
                </div>
              </div>

              {/* Settings Grid - 2 Columns */}
              <div className="grid grid-cols-1 md:grid-cols-2 gap-x-8 gap-y-8 border-t pt-8 border-dashed border-zinc-300 dark:border-zinc-800">
                
                {/* Column 1: Mode & Options */}
                <div className="space-y-8">
                  <div className="space-y-4">
                    <label className={cn("text-[10px] font-bold uppercase tracking-widest flex items-center gap-2", themeStyles.mutedText)}>
                      <Settings2 size={12} /> 核心模式
                    </label>
                    <div className={cn("flex flex-col p-1 rounded-2xl border", themeStyles.subBg, themeStyles.border)}>
                      <button onClick={() => setExportMode('4-col')} className={cn("w-full px-4 py-2.5 rounded-xl text-left text-[10px] font-bold transition-all flex items-center justify-between", exportMode === '4-col' ? (themeStyles.inputBg + " " + themeStyles.inputText + " shadow-sm") : (themeStyles.mutedText + " hover:opacity-80"))}>
                        双栏模式（仅提示）
                        {exportMode === '4-col' && <div className="w-1 h-1 rounded-full bg-blue-500 shadow-lg shadow-blue-500/50" />}
                      </button>
                      <button onClick={() => setExportMode('6-col')} className={cn("w-full px-4 py-2.5 rounded-xl text-left text-[10px] font-bold transition-all flex items-center justify-between", exportMode === '6-col' ? (themeStyles.inputBg + " " + themeStyles.inputText + " shadow-sm") : (themeStyles.mutedText + " hover:opacity-80"))}>
                        三栏模式（提示+词性）
                        {exportMode === '6-col' && <div className="w-1 h-1 rounded-full bg-blue-500 shadow-lg shadow-blue-500/50" />}
                      </button>
                      <button onClick={() => setExportMode('review')} className={cn("w-full px-4 py-2.5 rounded-xl text-left text-[10px] font-bold transition-all flex items-center justify-between", exportMode === 'review' ? (themeStyles.inputBg + " " + themeStyles.inputText + " shadow-sm") : (themeStyles.mutedText + " hover:opacity-80"))}>
                        三栏模式（复习资料）
                        {exportMode === 'review' && <div className="w-1 h-1 rounded-full bg-blue-500 shadow-lg shadow-blue-500/50" />}
                      </button>
                    </div>
                  </div>

                  <div className="space-y-4">
                    <label className={cn("text-[10px] font-bold uppercase tracking-widest flex items-center gap-2", themeStyles.mutedText)}>
                      <Sparkles size={12} /> 附加功能
                    </label>
                    <div className={cn("p-4 rounded-2xl border space-y-3", themeStyles.subBg, themeStyles.border)}>
                      <div className="flex items-center justify-between gap-4">
                        <span className={cn("text-[9px] font-bold uppercase", themeStyles.mutedText)}>综合测试</span>
                        <button 
                          onClick={() => setIsExamMode(!isExamMode)}
                          className={cn(
                            "w-8 h-4 rounded-full relative transition-colors border",
                            isExamMode 
                              ? (theme === 'light' ? "bg-black border-black" : theme === 'sepia' ? "bg-[#5B4636] border-[#5B4636]" : "bg-blue-600 border-blue-500")
                              : (themeStyles.toggleBg + " border-black/10")
                          )}
                        >
                          <div className={cn(
                            "absolute top-0.5 w-2.5 h-2.5 rounded-full transition-all",
                            isExamMode ? "left-4.5 bg-white" : "left-0.5 bg-gray-500"
                          )} />
                        </button>
                      </div>
                      <div className="flex items-center justify-between gap-4">
                        <span className={cn("text-[9px] font-bold uppercase", themeStyles.mutedText)}>自动跳转</span>
                        <button 
                          onClick={() => setEnterJump(!enterJump)}
                          className={cn(
                            "w-8 h-4 rounded-full relative transition-colors border",
                            enterJump 
                              ? (theme === 'light' ? "bg-black border-black" : theme === 'sepia' ? "bg-[#5B4636] border-[#5B4636]" : "bg-blue-600 border-blue-500")
                              : (themeStyles.toggleBg + " border-black/10")
                          )}
                        >
                          <div className={cn(
                            "absolute top-0.5 w-2.5 h-2.5 rounded-full transition-all",
                            enterJump ? "left-4.5 bg-white" : "left-0.5 bg-gray-500"
                          )} />
                        </button>
                      </div>
                      <div className="flex items-center justify-between gap-4">
                        <span className={cn("text-[9px] font-bold uppercase", themeStyles.mutedText)}>同步删除</span>
                        <button 
                          onClick={() => setIsSyncDelete(!isSyncDelete)}
                          className={cn(
                            "w-8 h-4 rounded-full relative transition-colors border",
                            isSyncDelete 
                              ? (theme === 'light' ? "bg-black border-black" : theme === 'sepia' ? "bg-[#5B4636] border-[#5B4636]" : "bg-blue-600 border-blue-500")
                              : (themeStyles.toggleBg + " border-black/10")
                          )}
                        >
                          <div className={cn(
                            "absolute top-0.5 w-2.5 h-2.5 rounded-full transition-all",
                            isSyncDelete ? "left-4.5 bg-white" : "left-0.5 bg-gray-500"
                          )} />
                        </button>
                      </div>

                      {theme === 'sepia' && (
                        <div className="flex flex-col gap-2 pt-2 border-t border-dashed border-black/10">
                          <div className="flex justify-between items-center">
                            <span className={cn("text-[9px] font-bold uppercase", themeStyles.mutedText)}>护眼强度</span>
                            <span className={cn("text-[10px] font-mono", themeStyles.inputText)}>{sepiaIntensity}%</span>
                          </div>
                          <input 
                            type="range" 
                            min="0" 
                            max="100" 
                            value={sepiaIntensity}
                            onChange={(e) => setSepiaIntensity(Number(e.target.value))}
                            className="w-full h-1 bg-black/10 rounded-full appearance-none cursor-pointer accent-[#5B4636]"
                          />
                        </div>
                      )}
                    </div>
                  </div>
                </div>

                {/* Column 2: Export & Preference */}
                <div className="space-y-8">
                  <div className="space-y-4">
                    <label className={cn("text-[10px] font-bold uppercase tracking-widest flex items-center gap-2", themeStyles.mutedText)}>
                      <Type size={12} /> 导出参数
                    </label>
                    <div className={cn("p-4 rounded-2xl border space-y-4", themeStyles.subBg, themeStyles.border)}>
                      <div className="flex flex-col gap-2">
                        <span className={cn("text-[9px] font-bold uppercase", themeStyles.mutedText)}>字体</span>
                        <select 
                          value={exportSettings.font}
                          onChange={(e) => setExportSettings({...exportSettings, font: e.target.value})}
                          className={cn(
                            "text-[10px] font-bold bg-transparent focus:outline-none border-b cursor-pointer", 
                            themeStyles.inputBorder, themeStyles.inputText
                          )}
                        >
                          <option value="Microsoft YaHei" className={themeStyles.inputBg + " " + themeStyles.inputText}>微软雅黑</option>
                          <option value="SimSun" className={themeStyles.inputBg + " " + themeStyles.inputText}>宋体</option>
                          <option value="SimHei" className={themeStyles.inputBg + " " + themeStyles.inputText}>黑体</option>
                          <option value="Arial" className={themeStyles.inputBg + " " + themeStyles.inputText}>Arial</option>
                          <option value="Times New Roman" className={themeStyles.inputBg + " " + themeStyles.inputText}>Times New Roman</option>
                        </select>
                      </div>
                      <div className="flex items-center justify-between gap-4">
                        <span className={cn("text-[9px] font-bold uppercase", themeStyles.mutedText)}>字号</span>
                        <input 
                          type="number" 
                          value={exportSettings.fontSize}
                          step="0.5"
                          min="1"
                          onChange={(e) => setExportSettings({...exportSettings, fontSize: Number(e.target.value)})}
                          className={cn("w-12 bg-transparent text-right font-bold text-[10px] focus:outline-none border-b", themeStyles.inputBorder, themeStyles.inputText)}
                        />
                      </div>
                      <div className="flex items-center justify-between gap-4">
                        <span className={cn("text-[9px] font-bold uppercase", themeStyles.mutedText)}>居中</span>
                        <button 
                          onClick={() => setExportSettings({...exportSettings, isCentered: !exportSettings.isCentered})}
                          className={cn(
                            "w-8 h-4 rounded-full relative transition-colors border",
                            exportSettings.isCentered 
                              ? (theme === 'light' ? "bg-black border-black" : theme === 'sepia' ? "bg-[#5B4636] border-[#5B4636]" : "bg-blue-600 border-blue-500")
                              : (themeStyles.toggleBg + " border-black/10")
                          )}
                        >
                          <div className={cn(
                            "absolute top-0.5 w-2.5 h-2.5 rounded-full transition-all",
                            exportSettings.isCentered ? "left-4.5 bg-white" : "left-0.5 bg-gray-500"
                          )} />
                        </button>
                      </div>
                    </div>
                  </div>

                  <div className="space-y-4">
                    <label className={cn("text-[10px] font-bold uppercase tracking-widest flex items-center gap-2", themeStyles.mutedText)}>
                      <Sparkles size={12} /> 系统偏好
                    </label>
                    <div className={cn("p-4 rounded-2xl border space-y-4", themeStyles.subBg, themeStyles.border)}>
                      <div className="flex flex-col gap-2">
                        <span className={cn("text-[9px] font-bold uppercase", themeStyles.mutedText)}>自动保存 (秒)</span>
                        <div className="flex items-center justify-end gap-2">
                          <input 
                            type="number"
                            min="0"
                            value={autoSaveInterval}
                            onChange={(e) => setAutoSaveInterval(Number(e.target.value))}
                            className={cn("bg-transparent text-right font-bold text-[10px] focus:outline-none border-b w-12", themeStyles.inputBorder, themeStyles.inputText)}
                          />
                          <span className="text-[8px] text-zinc-500 font-bold uppercase">{autoSaveInterval === 0 ? "关闭" : "S"}</span>
                        </div>
                      </div>
                    </div>
                  </div>
                </div>
              </div>

              {/* Independent Save Default Button */}
              <button 
                id="save-defaults-btn"
                onClick={saveDefaultSettings}
                className={cn(
                  "w-full py-4 rounded-2xl text-[10px] font-bold uppercase tracking-widest transition-all border shadow-lg hover:shadow-xl active:scale-[0.98]",
                  isSavingDefaults 
                    ? "bg-cyan-500 text-white border-cyan-500 shadow-cyan-500/20"
                    : theme === 'light' 
                    ? "bg-black text-white border-black hover:bg-zinc-800"
                    : theme === 'sepia'
                    ? "bg-[#5B4636] text-[#F4ECD8] border-[#5B4636] hover:bg-[#4A392C]"
                    : "bg-zinc-900 text-zinc-300 border-zinc-700 hover:bg-zinc-800 hover:text-white"
                )}
              >
                {isSavingDefaults ? "已保存默认配置 ✔" : "保存当前偏好为默认"}
              </button>
            </div>
          </aside>
        </main>

        <footer className={cn("mt-24 pt-10 border-t flex justify-between items-center text-[10px] font-bold tracking-widest opacity-30", theme !== 'dark' ? "border-black" : "border-zinc-800")}>
          <p>© 2026 Randomizer STUDIO</p>
        </footer>

        {showPreview && (
          <div className="fixed inset-0 z-50 flex flex-col p-4 md:p-12 bg-black/80 backdrop-blur-md animate-in fade-in duration-300">
            <div className={cn(
              "flex-1 flex flex-col max-w-6xl mx-auto w-full rounded-[48px] border shadow-2xl overflow-hidden",
              theme !== 'dark' ? "bg-[#F5F5F7] border-black" : "bg-[#111112] border-zinc-800"
            )}>
              {/* Header */}
              <div className={cn("p-6 md:px-12 md:py-8 border-b flex justify-between items-center", theme !== 'dark' ? "border-black/10" : "border-zinc-800")}>
                <div>
                  <h2 className="text-xl font-light tracking-tight uppercase">排版预览与交互</h2>
                  <p className="text-[10px] font-bold text-zinc-500 uppercase tracking-widest mt-1">支持拖拽排序与列宽调整</p>
                </div>
                <div className="flex gap-4">
                  <button 
                    onClick={() => handleExport(previewItems, previewTransformations, previewSentences)}
                    disabled={isExporting}
                    className={cn(
                      "px-8 py-3 rounded-2xl text-[11px] font-bold transition-all border flex items-center gap-2 shadow-sm",
                      theme !== 'dark' ? "bg-black text-white border-black hover:bg-zinc-800" : "bg-white text-black border-white hover:bg-zinc-100 font-bold"
                    )}
                  >
                    {isExporting ? <div className="w-4 h-4 border-2 border-zinc-200 border-t-zinc-800 rounded-full animate-spin" /> : <Download size={14} />} 导出 WORD
                  </button>
                  <button 
                    onClick={() => setShowPreview(false)}
                    className={cn(
                      "p-3 rounded-2xl transition-all border shadow-sm",
                      theme !== 'dark' ? "bg-white border-black hover:bg-gray-100 text-black" : "bg-zinc-900 border-zinc-800 hover:bg-zinc-800 text-white"
                    )}
                  >
                    <X size={20} />
                  </button>
                </div>
              </div>

              {/* Body */}
              <div className="flex-1 overflow-y-auto p-6 md:p-12 space-y-12 pretty-scrollbar backdrop-blur-sm">
                {/* Spacing Controls Bar */}
                <div className={cn(
                  "p-6 rounded-[32px] border flex flex-wrap items-center gap-8",
                  themeStyles.innerBg, themeStyles.border
                )}>
                  <div className="flex items-center gap-4">
                    <span className={cn("text-[10px] font-bold uppercase flex items-center gap-2", themeStyles.mutedText)}>
                       <Layout size={12} /> 行间距
                    </span>
                    <div className={cn("flex items-center border rounded-xl overflow-hidden px-3 py-1", themeStyles.innerBg, themeStyles.border)}>
                      <input 
                        type="number" 
                        step="0.1"
                        min="1"
                        value={exportSettings.lineSpacing}
                        onChange={(e) => setExportSettings({...exportSettings, lineSpacing: Number(e.target.value)})}
                        className={cn("w-12 bg-transparent text-center text-xs font-bold focus:outline-none", themeStyles.inputText)}
                      />
                    </div>
                  </div>

                  <div className="flex items-center gap-4">
                    <span className={cn("text-[10px] font-bold uppercase flex items-center gap-2", themeStyles.mutedText)}>
                       <GripVertical size={12} /> 列间距倍率
                    </span>
                    <div className={cn("flex items-center border rounded-xl overflow-hidden px-3 py-1", themeStyles.innerBg, themeStyles.border)}>
                      <input 
                        type="number" 
                        step="0.1"
                        min="0"
                        value={exportSettings.columnSpacing}
                        onChange={(e) => setExportSettings({...exportSettings, columnSpacing: Number(e.target.value)})}
                        className={cn("w-12 bg-transparent text-center text-xs font-bold focus:outline-none", themeStyles.inputText)}
                      />
                    </div>
                  </div>

                  <div className="h-4 w-px bg-zinc-400/20 md:block hidden" />
                  
                  <p className="text-[9px] font-medium text-zinc-400 italic">数值将实时同步至导出配置</p>
                </div>

                {/* Word Section */}
                {previewItems.length > 0 && (
                  <div className="space-y-6">
                    <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
                      <span className="text-[11px] font-bold uppercase tracking-widest px-4 py-2 bg-zinc-400/10 rounded-full w-fit">1. 单词列表 ({exportMode})</span>
                    </div>
                    
                  <DndContext sensors={sensors} collisionDetection={closestCenter} onDragEnd={handleDragEndItems}>
                    <SortableContext items={previewItems} strategy={verticalListSortingStrategy}>
                      <div 
                        ref={previewTableRef}
                        className={cn(
                          "grid gap-px border overflow-hidden rounded-xl relative",
                          theme !== 'dark' ? "border-black bg-black" : "border-zinc-700 bg-zinc-800"
                        )}
                        style={{
                          gridTemplateColumns: exportMode === '4-col' 
                            ? `${columnWidths.word}% ${columnWidths.pos}% ${columnWidths.word}% ${columnWidths.pos}%`
                            : exportMode === '6-col'
                            ? `${columnWidths.word/1.5}% ${columnWidths.pos/1.5}% ${columnWidths.word/1.5}% ${columnWidths.pos/1.5}% ${columnWidths.word/1.5}% ${columnWidths.pos/1.5}%`
                            : exportMode === 'review'
                            ? `${columnWidths.word * columnWidths.reviewRatio}% ${columnWidths.word * (1 - columnWidths.reviewRatio)}% ${columnWidths.pos}% ${columnWidths.word * columnWidths.reviewRatio}% ${columnWidths.word * (1 - columnWidths.reviewRatio)}% ${columnWidths.pos}%`
                            : `${columnWidths.word}% ${columnWidths.pos}% ${columnWidths.word}% ${columnWidths.pos}%`,
                          width: '100%'
                        }}
                      >
                        {/* Drag Handle Overlays */}
                        {exportMode === '4-col' && (
                          <>
                            <div 
                              onMouseDown={handleDividerDrag}
                              className="absolute top-0 bottom-0 w-1.5 cursor-col-resize z-30 hover:bg-blue-500/30 transition-colors"
                              style={{ left: `${columnWidths.word}%`, transform: 'translateX(-50%)' }}
                            />
                            <div 
                              onMouseDown={handleDividerDrag}
                              className="absolute top-0 bottom-0 w-1.5 cursor-col-resize z-30 hover:bg-blue-500/30 transition-colors"
                              style={{ left: `${50 + columnWidths.word}%`, transform: 'translateX(-50%)' }}
                            />
                          </>
                        )}
                        {exportMode === '6-col' && (
                          <>
                            <div 
                              onMouseDown={handleDividerDrag}
                              className="absolute top-0 bottom-0 w-1.5 cursor-col-resize z-30 hover:bg-blue-500/30 transition-colors"
                              style={{ left: `${columnWidths.word/1.5}%`, transform: 'translateX(-50%)' }}
                            />
                            <div 
                              onMouseDown={handleDividerDrag}
                              className="absolute top-0 bottom-0 w-1.5 cursor-col-resize z-30 hover:bg-blue-500/30 transition-colors"
                              style={{ left: `${33.333 + columnWidths.word/1.5}%`, transform: 'translateX(-50%)' }}
                            />
                            <div 
                              onMouseDown={handleDividerDrag}
                              className="absolute top-0 bottom-0 w-1.5 cursor-col-resize z-30 hover:bg-blue-500/30 transition-colors"
                              style={{ left: `${66.666 + columnWidths.word/1.5}%`, transform: 'translateX(-50%)' }}
                            />
                          </>
                        )}
                        {exportMode === 'review' && (
                          <>
                            <div 
                              onMouseDown={(e) => handleDividerDrag(e, 'review')}
                              className="absolute top-0 bottom-0 w-1.5 cursor-col-resize z-30 hover:bg-blue-500/30 transition-colors"
                              style={{ left: `${columnWidths.word * columnWidths.reviewRatio}%`, transform: 'translateX(-50%)' }}
                            />
                            <div 
                              onMouseDown={(e) => handleDividerDrag(e, 'main')}
                              className="absolute top-0 bottom-0 w-1.5 cursor-col-resize z-30 hover:bg-blue-500/30 transition-colors"
                              style={{ left: `${columnWidths.word}%`, transform: 'translateX(-50%)' }}
                            />
                            <div 
                              onMouseDown={(e) => handleDividerDrag(e, 'review')}
                              className="absolute top-0 bottom-0 w-1.5 cursor-col-resize z-30 hover:bg-blue-500/30 transition-colors"
                              style={{ left: `${50 + columnWidths.word * columnWidths.reviewRatio}%`, transform: 'translateX(-50%)' }}
                            />
                            <div 
                              onMouseDown={(e) => handleDividerDrag(e, 'main')}
                              className="absolute top-0 bottom-0 w-1.5 cursor-col-resize z-30 hover:bg-blue-500/30 transition-colors"
                              style={{ left: `${50 + columnWidths.word}%`, transform: 'translateX(-50%)' }}
                            />
                          </>
                        )}

                        {/* Header Row */}
                        <div className={cn("contents", theme !== 'dark' ? "text-black" : "text-white")}>
                          {exportMode === '4-col' ? (
                            <>
                              <div className={cn("p-3 text-[10px] font-bold text-center border-r border-current uppercase", theme !== 'dark' ? "bg-gray-100" : "bg-zinc-900")}>测试提示</div>
                              <div className={cn("p-3 text-[10px] font-bold text-center border-r border-current uppercase", theme !== 'dark' ? "bg-gray-100" : "bg-zinc-900")}>答案</div>
                              <div className={cn("p-3 text-[10px] font-bold text-center border-r border-current uppercase", theme !== 'dark' ? "bg-gray-100" : "bg-zinc-900")}>测试提示</div>
                              <div className={cn("p-3 text-[10px] font-bold text-center uppercase", theme !== 'dark' ? "bg-gray-100" : "bg-zinc-900")}>答案</div>
                            </>
                          ) : exportMode === '6-col' ? (
                            <>
                              <div className={cn("p-3 text-[10px] font-bold text-center border-r border-current uppercase", theme !== 'dark' ? "bg-gray-100" : "bg-zinc-900")}>测试提示</div>
                              <div className={cn("p-3 text-[10px] font-bold text-center border-r border-current uppercase", theme !== 'dark' ? "bg-gray-100" : "bg-zinc-900")}>词性</div>
                              <div className={cn("p-3 text-[10px] font-bold text-center border-r border-current uppercase", theme !== 'dark' ? "bg-gray-100" : "bg-zinc-900")}>测试提示</div>
                              <div className={cn("p-3 text-[10px] font-bold text-center border-r border-current uppercase", theme !== 'dark' ? "bg-gray-100" : "bg-zinc-900")}>词性</div>
                              <div className={cn("p-3 text-[10px] font-bold text-center border-r border-current uppercase", theme !== 'dark' ? "bg-gray-100" : "bg-zinc-900")}>测试提示</div>
                              <div className={cn("p-3 text-[10px] font-bold text-center uppercase", theme !== 'dark' ? "bg-gray-100" : "bg-zinc-900")}>词性</div>
                            </>
                          ) : (
                            <>
                              <div className={cn("p-3 text-[10px] font-bold text-center border-r border-current uppercase", theme !== 'dark' ? "bg-gray-100" : "bg-zinc-900")}>测试提示</div>
                              <div className={cn("p-3 text-[10px] font-bold text-center border-r border-current uppercase", theme !== 'dark' ? "bg-gray-100" : "bg-zinc-900")}>词性</div>
                              <div className={cn("p-3 text-[10px] font-bold text-center border-r border-current uppercase", theme !== 'dark' ? "bg-gray-100" : "bg-zinc-900")}>答案</div>
                              <div className={cn("p-3 text-[10px] font-bold text-center border-r border-current uppercase", theme !== 'dark' ? "bg-gray-100" : "bg-zinc-900")}>测试提示</div>
                              <div className={cn("p-3 text-[10px] font-bold text-center border-r border-current uppercase", theme !== 'dark' ? "bg-gray-100" : "bg-zinc-900")}>词性</div>
                              <div className={cn("p-3 text-[10px] font-bold text-center uppercase", theme !== 'dark' ? "bg-gray-100" : "bg-zinc-900")}>答案</div>
                            </>
                          )}
                        </div>

                        {/* Data Rows - Grouped by 2 or 3 to simulate Word table rows */}
                        {Array.from({ length: Math.ceil(previewItems.length / (exportMode === '6-col' ? 3 : 2)) }).map((_, rowIndex) => {
                          const itemsPerRow = exportMode === '6-col' ? 3 : 2;
                          const startIndex = rowIndex * itemsPerRow;
                          
                          return (
                            <div key={`row-${rowIndex}`} className="contents">
                              {Array.from({ length: itemsPerRow }).map((_, colIndex) => {
                                const itemIndex = startIndex + colIndex;
                                const item = previewItems[itemIndex];
                                const isLastInRow = colIndex === itemsPerRow - 1;

                                if (item) {
                                  return (
                                    <React.Fragment key={item.id}>
                                      <SortableItem id={item.id} className={cn("p-3 text-xs border-r border-current text-center break-all", theme !== 'dark' ? "bg-white" : "bg-zinc-950")}>
                                        <span className="opacity-80 font-serif">{item.term}</span>
                                      </SortableItem>
                                      {exportMode === '6-col' || exportMode === 'review' ? (
                                        <div className={cn("p-3 text-[10px] font-mono border-r border-current italic text-center text-zinc-500 break-all", theme !== 'dark' ? "bg-white" : "bg-zinc-950")}>
                                          {item.pos}
                                        </div>
                                      ) : null}
                                      {exportMode === '4-col' && (
                                        <div className={cn("p-3 text-xs text-center break-all", !isLastInRow ? "border-r border-current" : "", theme !== 'dark' ? "bg-white" : "bg-zinc-950")}>
                                          {item.pos}
                                        </div>
                                      )}
                                      {exportMode === 'review' && (
                                        <div className={cn("p-3 text-xs text-center break-all", !isLastInRow ? "border-r border-current" : "", theme !== 'dark' ? "bg-white" : "bg-zinc-950")}>
                                          {item.answer}
                                        </div>
                                      )}
                                      {exportMode === '6-col' && isLastInRow && (
                                        // Final border for 6-col which uses its own cells
                                        <div className="hidden" />
                                      )}
                                    </React.Fragment>
                                  );
                                } else {
                                  // Filler cells for uneven rows
                                  const cellsPerItem = exportMode === 'review' ? 3 : 2;
                                  return Array.from({ length: cellsPerItem }).map((_, fillerIdx) => (
                                    <div key={`filler-${itemIndex}-${fillerIdx}`} className={cn("p-3 min-h-[40px] border-current", (!isLastInRow || fillerIdx < cellsPerItem - 1) ? "border-r" : "", theme !== 'dark' ? "bg-white" : "bg-zinc-950")} />
                                  ));
                                }
                              })}
                            </div>
                          );
                        })}
                      </div>
                    </SortableContext>
                  </DndContext>
                  </div>
                )}

                {/* Transformations Section */}
                {isExamMode && previewTransformations.length > 0 && (
                  <div className="space-y-6">
                    <span className="text-[11px] font-bold uppercase tracking-widest px-4 py-2 bg-zinc-400/10 rounded-full">2. 变形考查</span>
                    <DndContext sensors={sensors} collisionDetection={closestCenter} onDragEnd={handleDragEndTrans}>
                      <SortableContext items={previewTransformations} strategy={verticalListSortingStrategy}>
                        <div className={cn(
                          "grid grid-cols-2 gap-px border rounded-xl overflow-hidden",
                          theme !== 'dark' ? "border-black bg-black" : "border-zinc-700 bg-zinc-800"
                        )}>
                          {previewTransformations.map((t) => (
                            <SortableItem key={t.id} id={t.id} className={cn("p-4 text-xs", theme !== 'dark' ? "bg-white" : "bg-zinc-950")}>
                               <div className="flex items-center gap-2">
                                <span className="opacity-70">{t.content}</span>
                                <div className="flex-1 border-b border-dashed opacity-30 mt-2"></div>
                               </div>
                            </SortableItem>
                          ))}
                        </div>
                      </SortableContext>
                    </DndContext>
                  </div>
                )}

                {/* Sentences Section */}
                {isExamMode && previewSentences.length > 0 && (
                  <div className="space-y-6">
                    <span className="text-[11px] font-bold uppercase tracking-widest px-4 py-2 bg-zinc-400/10 rounded-full">3. 句子考查</span>
                    <DndContext sensors={sensors} collisionDetection={closestCenter} onDragEnd={handleDragEndSent}>
                      <SortableContext items={previewSentences} strategy={verticalListSortingStrategy}>
                        <div className="space-y-4">
                          {previewSentences.map((s, idx) => (
                            <SortableItem key={s.id} id={s.id} className={cn("p-6 rounded-3xl border group transition-all hover:border-black/40", theme !== 'dark' ? "bg-white border-black/10" : "bg-zinc-900/40 border-zinc-800/50")}>
                              <div className="flex gap-4">
                                <span className="text-xs font-mono font-bold opacity-30">#{(idx + 1).toString().padStart(2, '0')}</span>
                                <div className="space-y-4 flex-1">
                                  <p className="text-sm font-medium leading-relaxed">{s.content.replace(/^\d+[\.\s、]+/, '').trim()}</p>
                                  <div className="w-full h-px bg-current opacity-10"></div>
                                </div>
                              </div>
                            </SortableItem>
                          ))}
                        </div>
                      </SortableContext>
                    </DndContext>
                  </div>
                )}
                
                <div className="h-24"></div>
              </div>
              
              {/* Footer */}
              <div className={cn("p-6 md:px-12 md:py-8 border-t flex items-center justify-between", theme !== 'dark' ? "bg-white border-black/10" : "bg-[#0A0A0B] border-zinc-800")}>
                <p className="text-[10px] font-bold uppercase text-zinc-500 flex items-center gap-2 italic">
                  <Layout size={12} /> 排版遵循 Word 标准渲染逻辑。调整结束后点击上方导出按钮。
                </p>
                <div className="flex gap-3">
                   <div className="flex items-center gap-2 px-3 py-1.5 rounded-full bg-green-500/10 border border-green-500/20">
                      <div className="w-2 h-2 rounded-full bg-green-500" />
                      <span className="text-[9px] font-bold text-green-600 uppercase">实时同步已开启</span>
                   </div>
                </div>
              </div>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}
