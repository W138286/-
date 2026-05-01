/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

/**
 * App.tsx - v5.2 PRO
 * Updated at: 2026-04-28
 */
import React, { useState, useCallback, useRef, useEffect, useMemo } from 'react';
import { motion, AnimatePresence } from 'motion/react';
import { 
  Download, 
  Shuffle, 
  Trash2,
  FileText,
  Settings2,
  Sparkles,
  Sun,
  Moon,
  Type,
  X,
  Layout,
  GripVertical,
  Copy,
  ChevronRight,
  Eye,
  Monitor
} from 'lucide-react';
import { 
  Document, 
  Packer, 
  Table, 
  TableRow, 
  TableCell, 
  TableLayoutType,
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
  MouseSensor,
  TouchSensor,
  useSensor,
  useSensors,
  DragEndEvent
} from '@dnd-kit/core';
import {
  arrayMove,
  SortableContext,
  sortableKeyboardCoordinates,
  rectSortingStrategy,
  verticalListSortingStrategy,
  useSortable
} from '@dnd-kit/sortable';
import { CSS } from '@dnd-kit/utilities';

// --- Components ---
function SortableItem({ id, children, className, style: customStyle }: { id: string, children: React.ReactNode, className?: string, key?: any, style?: React.CSSProperties }) {
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
    ...customStyle
  };

  return (
    <div ref={setNodeRef} style={style} className={cn("relative group", className)}>
      <div {...attributes} {...listeners} className="absolute left-1 top-1/2 -translate-y-1/2 opacity-0 group-hover:opacity-100 cursor-grab active:cursor-grabbing transition-opacity z-[50] p-0.5 rounded-sm bg-zinc-500/10 backdrop-blur-sm">
        <GripVertical size={14} className="text-zinc-500" />
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
  onKeyDown,
  theme = 'light',
  fontSize = 14
}: {
  value: string;
  onChange: (e: React.ChangeEvent<HTMLTextAreaElement>) => void;
  onScroll?: (e: React.UIEvent<HTMLTextAreaElement>) => void;
  onKeyDown?: (e: React.KeyboardEvent<HTMLTextAreaElement>) => void;
  placeholder?: string;
  className?: string;
  themeStyles: any;
  textareaRef?: React.RefObject<HTMLTextAreaElement | null>;
  theme?: 'light' | 'dark' | 'sepia';
  fontSize?: number;
}) {
  const lineColor = theme === 'dark' ? 'rgba(255, 255, 255, 0.25)' : theme === 'sepia' ? 'rgba(91, 70, 54, 0.45)' : 'rgba(0, 0, 0, 0.2)';
  
  const lineHeight = Math.max(20, Math.round(fontSize * 1.857));

  const handleMouseDown = (e: React.MouseEvent<HTMLElement>) => {
    if (e.button !== 0) return; // Only process left click
    
    if (textareaRef?.current) {
      if (e.target !== textareaRef.current) {
        e.preventDefault();
        textareaRef.current.focus();
        
        // Place cursor at the end of the content
        const len = textareaRef.current.value.length;
        textareaRef.current.setSelectionRange(len, len);
      }
    }
  };

  return (
    <div 
      className="flex-1 flex cursor-text min-h-full relative"
      onMouseDown={handleMouseDown}
    >
      <textarea
        ref={textareaRef}
        value={value}
        onChange={onChange}
        onScroll={onScroll}
        onKeyDown={onKeyDown}
        onMouseDown={handleMouseDown}
        placeholder={placeholder}
        spellCheck={false}
        wrap="off"
        className={cn(
          "block m-0 border-0 flex-1 px-6 py-6 focus:outline-none resize-none font-sans transition-colors bg-transparent min-h-full w-full",
          themeStyles.inputText,
          className
        )}
        style={{ 
          fontSize: `${fontSize}px`,
          lineHeight: `${lineHeight}px`,
          backgroundImage: `repeating-linear-gradient(transparent, transparent ${lineHeight - 1}px, ${lineColor} ${lineHeight - 1}px, ${lineColor} ${lineHeight}px)`,
          backgroundSize: `100% ${lineHeight}px`,
          backgroundPosition: '0 24px',
          backgroundAttachment: 'local'
        }}
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
  state = { hasError: false, error: null as Error | null };

  constructor(props: {children: React.ReactNode}) {
    super(props);
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
    // @ts-ignore
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
  
  // Maintain equal width layout whenever mode changes (as user requested equal split as default)
  useEffect(() => {
    if (exportMode === '4-col') {
      setColumnWidths({
        word: 25, pos: 25, reviewRatio: 0.5,
        leftSplit: 50,
        wordRight: 25, posRight: 25, reviewRatioRight: 0.5
      });
      setInputColRatio({ word: 50, pos: 50, ans: 0 }); // Two equal cols
    } else {
      setColumnWidths({
        word: 33.333, pos: 16.667, reviewRatio: 0.5,
        leftSplit: 50,
        wordRight: 33.333, posRight: 16.667, reviewRatioRight: 0.5
      });
      setInputColRatio({ word: 33.33, pos: 33.33, ans: 33.34 }); // Three equal cols
    }
  }, [exportMode]);
  
  const [inputColRatio, setInputColRatio] = useState<{word: number, pos: number, ans: number}>({ word: 33.33, pos: 33.33, ans: 33.34 });

  const [inputActiveDivider, setInputActiveDivider] = useState<'left' | 'right' | null>(null);

  useEffect(() => {
    const handleMouseMove = (e: MouseEvent) => {
      if (!inputActiveDivider) return;
      const container = document.getElementById('input-area-container');
      if (!container) return;
      
      const rect = container.getBoundingClientRect();
      const x = e.clientX - rect.left;
      const percentage = Math.max(10, Math.min(90, (x / rect.width) * 100)); // Limit to 10% - 90%
      
      setInputColRatio(prev => {
        if (exportMode === '6-col' || exportMode === '4-col') {
          const total = prev.word + prev.pos;
          const newWord = (percentage / 100) * total;
          const newPos = total - newWord;
          return {
            ...prev,
            word: Math.max(10, newWord),
            pos: Math.max(10, newPos)
          };
        } else {
          if (inputActiveDivider === 'left') {
            const newPos = prev.word + prev.pos - percentage;
            if (newPos < 10) return prev;
            return {
              ...prev,
              word: percentage,
              pos: newPos
            };
          } else {
            const newPos = percentage - prev.word;
            if (newPos < 10) return prev;
            const newAns = 100 - percentage;
            if (newAns < 10) return prev;
            return {
              ...prev,
              pos: newPos,
              ans: newAns
            };
          }
        }
      });
    };

    const handleMouseUp = () => {
      if (inputActiveDivider) {
        setInputActiveDivider(null);
      }
    };

    if (inputActiveDivider) {
      window.addEventListener('mousemove', handleMouseMove);
      window.addEventListener('mouseup', handleMouseUp);
    }
    return () => {
      window.removeEventListener('mousemove', handleMouseMove);
      window.removeEventListener('mouseup', handleMouseUp);
    };
  }, [inputActiveDivider, exportMode, inputColRatio]);
  const [isExamMode, setIsExamMode] = useState<boolean>(() => {
    if (typeof window !== 'undefined') {
      return storage.getItem('app-exam-mode') === 'true';
    }
    return false;
  });
  const [examType, setExamType] = useState<'EN' | 'CN'>('EN');
  const [isExporting, setIsExporting] = useState(false);
  const [showPreview, setShowPreview] = useState(false);
  const [showStylePicker, setShowStylePicker] = useState(false);
  const [exportStyle, setExportStyle] = useState<'table' | 'text'>('table');
  const [isDragging, setIsDragging] = useState(false);
  const [previewItems, setPreviewItems] = useState<{ id: string; term: string; pos: string; answer: string }[]>([]);
  const [previewTransformations, setPreviewTransformations] = useState<{ id: string; content: string }[]>([]);
  const [previewSentences, setPreviewSentences] = useState<{ id: string; content: string }[]>([]);
  const [columnWidths, setColumnWidths] = useState<{ word: number; pos: number; reviewRatio: number; leftSplit: number; wordRight: number; posRight: number; reviewRatioRight: number }>({ word: 25, pos: 25, reviewRatio: 0.5, leftSplit: 50, wordRight: 25, posRight: 25, reviewRatioRight: 0.5 });
  const [activeDivider, setActiveDivider] = useState<{ type: string, side: string } | null>(null);
  const [isSymmetricSync, setIsSymmetricSync] = useState(() => {
    if (typeof window !== 'undefined') {
      const saved = storage.getItem('app-symmetric-sync');
      return saved !== null ? saved === 'true' : true;
    }
    return true;
  });
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

  const [enterNewLine, setEnterNewLine] = useState<boolean>(() => {
    if (typeof window !== 'undefined') {
      const saved = storage.getItem('app-enter-newline');
      return saved !== null ? saved === 'true' : false;
    }
    return false;
  });

  const [inputFontSize, setInputFontSize] = useState<number>(() => {
    if (typeof window !== 'undefined') {
      const saved = storage.getItem('app-input-font-size');
      return saved ? Number(saved) : 14;
    }
    return 14;
  });

  const [sepiaIntensity, setSepiaIntensity] = useState<number>(() => {
    if (typeof window !== 'undefined') {
      const saved = storage.getItem('app-sepia-intensity');
      return saved ? Number(saved) : 50;
    }
    return 50;
  });

  const [brightness, setBrightness] = useState<number>(() => {
    if (typeof window !== 'undefined') {
      const saved = storage.getItem('app-brightness');
      return saved ? Number(saved) : 100;
    }
    return 100;
  });

  // Memoized line arrays to prevent splitting on every render for large files
  const inputLineArray = useMemo(() => Array.from({ length: input.split('\n').length }), [input]);
  const posLineArray = useMemo(() => Array.from({ length: posInput.split('\n').length }), [posInput]);
  const answerLineArray = useMemo(() => Array.from({ length: answerInput.split('\n').length }), [answerInput]);

  useEffect(() => {
    storage.setItem('app-enter-jump', String(enterJump));
  }, [enterJump]);

  useEffect(() => {
    storage.setItem('app-enter-newline', String(enterNewLine));
  }, [enterNewLine]);

  useEffect(() => {
    storage.setItem('app-brightness', String(brightness));
  }, [brightness]);

  useEffect(() => {
    storage.setItem('app-input-font-size', String(inputFontSize));
    // Force sync of scroll position when font size changes because textarea scroll anchoring might shift scrollTop without firing onScroll
    requestAnimationFrame(() => {
      if (wordInputRef.current) {
        const top = wordInputRef.current.scrollTop;
        if (scrollRef.current) scrollRef.current.scrollTop = top;
        if (posInputRef.current) posInputRef.current.scrollTop = top;
        if (ansInputRef.current) ansInputRef.current.scrollTop = top;
      }
    });
  }, [inputFontSize]);

  useEffect(() => {
    if (showPreview) {
      document.body.style.overflow = 'hidden';
    } else {
      document.body.style.overflow = 'auto';
    }
  }, [showPreview]);

  const themeStyles = useMemo(() => {
    const intensity = sepiaIntensity / 100;
    
    // Sepia calculation: interpolate between light sepia and deep sepia
    // Base: #F4ECD8 (244, 236, 216) -> Deep: #D2C2A4 (210, 194, 164)
    const r = Math.round(244 - (244 - 210) * intensity);
    const g = Math.round(236 - (236 - 194) * intensity);
    const b = Math.round(216 - (216 - 164) * intensity);
    const sepiaBgColor = `rgb(${r}, ${g}, ${b})`;

    // Sub pattern for sepia (darker than bg) 
    // Base: #EBE0C5 (235, 224, 197) -> Deep: #C4B495 (196, 180, 149)
    const subR = Math.round(235 - (235 - 196) * intensity);
    const subG = Math.round(224 - (224 - 180) * intensity);
    const subB = Math.round(197 - (197 - 149) * intensity);
    const sepiaSubBgColor = `rgb(${subR}, ${subG}, ${subB})`;
    const sepiaInnerBgColor = `rgba(${subR}, ${subG}, ${subB}, 0.4)`;

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
        borderMuted: "border-zinc-700",
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
        customSubBg: sepiaSubBgColor,
        customInnerBg: sepiaInnerBgColor,
        border: "border-[#D2C2A4]",
        borderMuted: "border-[#D2C2A4]/50",
        subBg: "", // Use inline style with customSubBg instead
        innerBg: "", // Use inline style with customInnerBg instead
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
    // 更加稳健的分割和初步过滤
    const lines = text.split(/\r?\n/).map(line => line.trim()).filter(line => line.length > 0);
    const entries: { word: string; pos: string; answer: string }[] = [];

    lines.forEach(line => {
      // 优化正则：支持复合词性（如 n/v.）并允许 / \ | 作为前缀或连接符
      const singlePosPattern = `(?:${POS_LIST.join('|')}|phr|phrase)`;
      const multiPosPattern = `${singlePosPattern}(?:[\\s/\\\\|]+${singlePosPattern})*`;
      const posRegex = new RegExp(`(^|\\s|/|\\\\|\\|)\\s*(${multiPosPattern})(?:\\.?)(?=\\s|[^a-zA-Z]|$)`, 'i');
      const match = line.match(posRegex);
      
      let word = "";
      let pos = "";
      let answer = "";

      if (match) {
        const posIndex = match.index!;
        pos = match[2].toLowerCase();
        if (pos.includes('phr') || pos.includes('phrase')) {
          pos = pos.replace(/phrase/g, 'phr');
        }
        
        // 规范化词性格式
        if (!pos.endsWith('.')) {
          pos += '.';
        }
        
        word = line.substring(0, posIndex).trim();
        if (match[1] && word.endsWith(match[1].trim())) {
          word = word.substring(0, word.length - match[1].trim().length).trim();
        }
        
        answer = line.substring(posIndex + match[0].length).trim();
      } else {
        // 如果整行只是一个斜杠或由斜杠组成，将其视为“空单词，词性为/”
        if (/^[\\/|\s]+$/.test(line)) {
          word = "";
          pos = line.trim();
        } else {
          // 查找第一个非 ASCII 字符作为答案开头
          const chineseMatch = line.match(/[^\x00-\x7F\u3001\uff3c/|]/);
          let rawWordPart = line;

          if (chineseMatch) {
            const index = chineseMatch.index!;
            rawWordPart = line.substring(0, index).trim();
            answer = line.substring(index).trim();
          }

          // 如果单词部分最后以 / \ | 等符号结尾，将其提取到词性
          const slashMatch = rawWordPart.match(/^(.*?)[\s]*([\\/|]+)[\s]*$/);
          if (slashMatch) {
            word = slashMatch[1].trim();
            pos = slashMatch[2].trim();
          } else {
            word = rawWordPart.trim();
          }
        }
      }

      // 合并逻辑：将分行显示的单词、词性和答案合并到同一个条目
      if (entries.length > 0) {
        const last = entries[entries.length - 1];
        // 如果当前行没有单词但有词性或答案，且上一个条目只有单词，则合并
        if (!word && (pos || answer) && last.word && !last.pos && !last.answer) {
          last.pos = pos;
          last.answer = answer;
          return;
        }
        // 如果当前行只有答案，且上一个条目已经有单词和词性但没答案，则合并
        if (!word && !pos && answer && last.word && last.pos && !last.answer) {
          last.answer = answer;
          return;
        }
      }

      entries.push({ word, pos, answer });
    });

    const parsedWords = entries.map(e => e.word);
    const parsedPos = entries.map(e => e.pos);
    const parsedAnswers = entries.map(e => e.answer);

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

  const getVisibleCols = () => {
    if (exportMode === 'review') return 3;
    if (exportMode === '6-col') return 2;
    return 1;
  };

  const handleArrowJump = (
    e: React.KeyboardEvent<HTMLTextAreaElement>,
    currentCol: 1 | 2 | 3
  ) => {
    const visibleCols = getVisibleCols();
    
    if (e.key === 'ArrowLeft') {
      const { selectionStart, selectionEnd, value } = e.currentTarget;
      const info = getLineInfo(value, selectionStart);

      if (info.isAtStart && selectionStart === selectionEnd) {
        let targetCol: 1 | 2 | 3;
        let targetLine: number;

        if (currentCol === 1) {
          if (info.lineIndex === 0) return;
          if (visibleCols === 1) {
            targetCol = 1;
          } else {
            targetCol = visibleCols as 1 | 2 | 3;
          }
          targetLine = info.lineIndex - 1;
        } else {
          targetCol = (currentCol - 1) as 1 | 2 | 3;
          targetLine = info.lineIndex;
        }

        e.preventDefault();
        const targetRef = targetCol === 1 ? wordInputRef : (targetCol === 2 ? posInputRef : ansInputRef);
        setTimeout(() => {
          if (targetRef.current) {
            targetRef.current.focus();
            const pos = getPosAtLineEdge(targetRef.current.value, targetLine, true);
            targetRef.current.setSelectionRange(pos, pos);
          }
        }, 30);
      }
    } else if (e.key === 'ArrowRight') {
      const { selectionStart, selectionEnd, value } = e.currentTarget;
      const info = getLineInfo(value, selectionStart);
      
      if (info.isAtEnd && selectionStart === selectionEnd) {
        let targetCol: 1 | 2 | 3;
        let targetLine: number;
        
        if (currentCol === visibleCols) {
          targetCol = 1;
          targetLine = info.lineIndex + 1;
          if (targetLine >= info.totalLines) return;
        } else {
          targetCol = (currentCol + 1) as 1 | 2 | 3;
          targetLine = info.lineIndex;
        }

        e.preventDefault();
        const targetRef = targetCol === 1 ? wordInputRef : (targetCol === 2 ? posInputRef : ansInputRef);
        if (targetRef.current) {
          targetRef.current.focus();
          setTimeout(() => {
            if (targetRef.current) {
              const pos = getPosAtLineEdge(targetRef.current.value, targetLine, false);
              targetRef.current.setSelectionRange(pos, pos);
            }
          }, 0);
        }
      }
    }
  };

  const handleSyncDeletions = (
    ref: React.RefObject<HTMLTextAreaElement>,
    oldVal: string, 
    newVal: string, 
    setSelf: React.Dispatch<React.SetStateAction<string>>, 
    others: {val: string, setter: React.Dispatch<React.SetStateAction<string>>}[]
  ) => {
    let finalSelf = newVal;
    let finalOthers = others.map(o => o.val);

    if (isSyncDelete) {
      const oldLines = oldVal.split('\n');
      let newLines = newVal.split('\n');
      let deletedIndices: number[] = [];

      if (newLines.length < oldLines.length) {
        let start = 0;
        while (start < newLines.length && oldLines[start] === newLines[start]) {
          start++;
        }
        
        let endOld = oldLines.length - 1;
        let endNew = newLines.length - 1;
        
        while (endNew >= start && endOld >= start && oldLines[endOld] === newLines[endNew]) {
          endOld--;
          endNew--;
        }
        
        for (let i = start; i <= endOld; i++) {
          deletedIndices.push(i);
        }
      } else if (newLines.length === oldLines.length) {
        for (let i = 0; i < newLines.length; i++) {
          if (newLines[i] === '' && oldLines[i] !== '') {
            deletedIndices.push(i);
          }
        }
      }

      if (deletedIndices.length > 0) {
        if (newLines.length === oldLines.length) {
          newLines = newLines.filter((_, i) => !deletedIndices.includes(i));
        }
        finalSelf = newLines.join('\n');
        
        finalOthers = others.map(o => {
          const oLines = o.val.split('\n');
          return oLines.filter((_, i) => !deletedIndices.includes(i)).join('\n');
        });
        
        if (ref.current) {
          const targetLine = Math.max(0, deletedIndices[0] - 1);
          const pos = getPosAtLineEdge(finalSelf, targetLine, true);
          setTimeout(() => {
            if (ref.current) {
              ref.current.focus();
              ref.current.setSelectionRange(pos, pos);
            }
          }, 0);
        }
      }
    }

    // Pad lines synchronously to ensure all columns have the same number of lines
    const maxLines = Math.max(
      finalSelf.split('\n').length,
      ...finalOthers.map(o => o.split('\n').length)
    );

    const padLines = (val: string, count: number) => {
      const currentLines = val.split('\n');
      if (currentLines.length < count) {
        return val + '\n'.repeat(count - currentLines.length);
      }
      return val;
    };

    const paddedSelf = padLines(finalSelf, maxLines);
    finalOthers = finalOthers.map(val => padLines(val, maxLines));

    setSelf(paddedSelf);
    others.forEach((o, i) => o.setter(finalOthers[i]));

    // Restore cursor if the current input was padded asynchronously/synchronously and changed the text length without user intervention
    if (paddedSelf !== newVal && ref.current && !isSyncDelete) {
      // If we padded, user just deleted something but we added back newlines,
      // or user typed normally and we padded.
      // Typically, native backspace changes focus position correctly. But our padding might disrupt it.
      const selStart = ref.current.selectionStart;
      const selEnd = ref.current.selectionEnd;
      requestAnimationFrame(() => {
        setTimeout(() => {
          if (ref.current) {
            ref.current.setSelectionRange(selStart, selEnd);
          }
        }, 10);
      });
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
    useSensor(MouseSensor, {
      activationConstraint: {
        delay: 250,
        tolerance: 5,
      },
    }),
    useSensor(TouchSensor, {
      activationConstraint: {
        delay: 250,
        tolerance: 5,
      },
    }),
    useSensor(KeyboardSensor, {
      coordinateGetter: sortableKeyboardCoordinates,
    })
  );

  // Sync line counts across all columns to ensure alignment and clickability - moved to synchronous processing in handleSyncDeletions
  
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
    setShowStylePicker(true);
  };

  const startPreview = (style: 'table' | 'text') => {
    setExportStyle(style);
    setShowStylePicker(false);

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
    
    if (exportMode === '4-col') {
      setColumnWidths({
        word: 25, pos: 25, reviewRatio: 0.5,
        leftSplit: 50,
        wordRight: 25, posRight: 25, reviewRatioRight: 0.5
      });
    } else {
      setColumnWidths({
        word: 33.333, pos: 16.667, reviewRatio: 0.5,
        leftSplit: 50,
        wordRight: 33.333, posRight: 16.667, reviewRatioRight: 0.5
      });
    }
    
    setShowPreview(true);
  };

  const handleEnterJump = (
    e: React.KeyboardEvent<HTMLTextAreaElement>,
    currentCol: 1 | 2 | 3
  ) => {
    if (e.key === 'Enter' && !e.shiftKey) {
      if (!enterJump && !enterNewLine) {
        // According to user request: if enterJump is OFF, Enter jumps to next line in the same column.
        // We will process this normally below.
      }
      e.preventDefault();
      
      const textarea = e.currentTarget;
      const { lineIndex } = getLineInfo(textarea.value, textarea.selectionStart);
      const visibleCols = getVisibleCols();
      const isLastColumnOfMode = currentCol === visibleCols;
      
      let targetLine = lineIndex;
      let targetCol = currentCol;
      let isMovingToNextRow = false;

      if (enterJump) {
        if (isLastColumnOfMode) {
          targetCol = 1;
          targetLine += 1;
          isMovingToNextRow = true;
        } else {
          targetCol = (currentCol + 1) as 1 | 2 | 3;
        }
      } else {
        // enterJump is OFF: native-like jump is to the next line in the EXACT SAME column
        targetCol = currentCol;
        targetLine += 1;
        isMovingToNextRow = true;
      }

      let newWord = input;
      let newPos = posInput;
      let newAns = answerInput;
      let atEnd = true;

      if (isMovingToNextRow) {
        if (enterNewLine) {
          // "当开启时，若换行到了下一行，则在下一行空出一行，下方所有的内容都往下移一行"
          const insertEmptyLine = (val: string) => {
            const linesList = val.split('\n');
            while (linesList.length < targetLine) {
              linesList.push('');
            }
            linesList.splice(targetLine, 0, ''); 
            return linesList.join('\n');
          };
          newWord = insertEmptyLine(input);
          newPos = insertEmptyLine(posInput);
          newAns = insertEmptyLine(answerInput);
          atEnd = false; // "光标在该行最开头位置"
        } else {
          // "自动换行功能默认关闭。但关闭时，换行按照现有逻辑运动。附加一条换行时优先级最后的指令：若换行的位置已有内容，自动移动到该行该列内容的最后一位。"
          const padLines = (val: string) => {
            const linesList = val.split('\n');
            while (linesList.length <= targetLine) {
              linesList.push('');
            }
            return linesList.join('\n');
          };
          newWord = padLines(input);
          newPos = padLines(posInput);
          newAns = padLines(answerInput);
          atEnd = true; 
        }
      }

      // Sync state if it changed
      if (newWord !== input) setInput(newWord);
      if (newPos !== posInput) setPosInput(newPos);
      if (newAns !== answerInput) setAnswerInput(newAns);

      const targetVal = targetCol === 1 ? newWord : (targetCol === 2 ? newPos : newAns);
      const targetRef = targetCol === 1 ? wordInputRef : (targetCol === 2 ? posInputRef : ansInputRef);
      
      const pos = getPosAtLineEdge(targetVal, targetLine, atEnd);
      const prevScrollTop = textarea.scrollTop;
      const clientHeight = textarea.clientHeight;

      setTimeout(() => {
        if (targetRef.current) {
          targetRef.current.focus();
          targetRef.current.setSelectionRange(pos, pos);
          
          // Scroll to ensure visibility dynamically
          const lineTop = targetLine * 26; // approx line height
          const lineBottom = lineTop + 26;
          if (lineBottom > prevScrollTop + clientHeight) {
            targetRef.current.scrollTop = lineBottom - clientHeight;
          } else if (lineTop < prevScrollTop) {
            targetRef.current.scrollTop = lineTop;
          } else {
            targetRef.current.scrollTop = prevScrollTop;
          }
        }
      }, 0);
    }
  };

  const handleBackspaceJump = (
    e: React.KeyboardEvent<HTMLTextAreaElement>,
    currentCol: 1 | 2 | 3
  ) => {
    if (e.key === 'Backspace' && !e.shiftKey) {
      const { lineIndex } = getLineInfo(e.currentTarget.value, e.currentTarget.selectionStart);
      const lines = e.currentTarget.value.split('\n');
      const currentLineContent = lines[lineIndex] || '';
      
      // Only jump if current line in this column is empty and cursor is at start
      if (currentLineContent === '' && e.currentTarget.selectionStart === e.currentTarget.selectionEnd) {
        if (!enterJump) return;

        const prevScrollTop = e.currentTarget.scrollTop;
        let targetCol: 1 | 2 | 3;
        let targetLine = lineIndex;
        const visibleCols = getVisibleCols();

        if (currentCol === 1) {
          if (lineIndex === 0) return; // Top left
          targetCol = visibleCols as 1 | 2 | 3;
          targetLine = lineIndex - 1;
          
          e.preventDefault();
          
          // Delete the current line across all inputs if it's the first column
          const removeLineAt = (val: string, index: number) => {
            const l = val.split('\n');
            if (index < l.length) {
              l.splice(index, 1);
            }
            return l.join('\n');
          };
          
          setInput(prev => removeLineAt(prev, lineIndex));
          setPosInput(prev => removeLineAt(prev, lineIndex));
          setAnswerInput(prev => removeLineAt(prev, lineIndex));
        } else {
          targetCol = (currentCol - 1) as 1 | 2 | 3;
          e.preventDefault();
        }

        const targetRef = targetCol === 1 ? wordInputRef : (targetCol === 2 ? posInputRef : ansInputRef);
        setTimeout(() => {
          if (targetRef.current) {
            targetRef.current.focus();
            const currentVal = targetRef.current.value;
            const pos = getPosAtLineEdge(currentVal, targetLine, true);
            targetRef.current.setSelectionRange(pos, pos);
            // Restore scroll position to prevent jumping to the bottom
            targetRef.current.scrollTop = prevScrollTop;
          }
        }, 50);
      }
    }
  };

  const handleDividerDrag = (e: React.MouseEvent, type: 'main' | 'review' | 'split' = 'main', side: 'left' | 'right' | 'center' = 'center') => {
    e.preventDefault();
    e.stopPropagation();
    if (!previewTableRef.current) return;
    
    const container = previewTableRef.current;
    const startX = e.clientX;
    const cw = { ...columnWidths };
    const containerWidth = container.offsetWidth;

    if (containerWidth <= 0) return;

    setActiveDivider({ type, side });
    document.body.style.userSelect = 'none';
    document.body.style.cursor = 'col-resize';

    const onMouseMove = (moveEvent: MouseEvent) => {
      const deltaX = moveEvent.clientX - startX;
      const deltaPercent = (deltaX / containerWidth) * 100;
      
      setColumnWidths(prev => {
        let next = { ...prev };

        if (type === 'split') {
          next.leftSplit = Math.max(10, Math.min(90, cw.leftSplit + deltaPercent));
          return next;
        }

        if (type === 'review') {
          if (side === 'left') {
            let change = (deltaPercent * 100) / (cw.leftSplit * cw.word / 100);
            let newRatio = cw.reviewRatio + change / 100;
            newRatio = Math.max(0.1, Math.min(0.9, newRatio));
            next.reviewRatio = newRatio;
            if (isSymmetricSync) {
              next.reviewRatioRight = newRatio;
            }
          } else if (side === 'right') {
            let changeR = (deltaPercent * 100) / ((100 - cw.leftSplit) * cw.wordRight / 100);
            let newRatioR = cw.reviewRatioRight + changeR / 100;
            newRatioR = Math.max(0.1, Math.min(0.9, newRatioR));
            next.reviewRatioRight = newRatioR;
            if (isSymmetricSync) {
              next.reviewRatio = newRatioR;
            }
          }
          return next;
        }

        if (type === 'main') {
          if (side === 'left') {
            if (exportMode === '4-col') {
              let dword = (deltaPercent * 100) / cw.leftSplit;
              let newWordWidth = cw.word + dword;
              newWordWidth = Math.max(5, Math.min(45, newWordWidth));
              next.word = newWordWidth;
              next.pos = 50 - newWordWidth;
            } else {
              let dpos = (deltaPercent * 100) / cw.leftSplit;
              let newWordWidth = cw.word - dpos;
              newWordWidth = Math.max(5, Math.min(45, newWordWidth));
              
              let newReviewRatio = (cw.word * cw.reviewRatio) / newWordWidth;
              
              next.word = newWordWidth;
              next.pos = 50 - newWordWidth;
              next.reviewRatio = Math.max(0.1, Math.min(0.9, newReviewRatio));
            }

            if (isSymmetricSync) {
              next.wordRight = next.word;
              next.posRight = next.pos;
              next.reviewRatioRight = next.reviewRatio;
            }
          } else if (side === 'right') {
            if (exportMode === '4-col') {
              let dwordR = (deltaPercent * 100) / (100 - cw.leftSplit);
              let newWordWidthR = cw.wordRight + dwordR;
              newWordWidthR = Math.max(5, Math.min(45, newWordWidthR));
              next.wordRight = newWordWidthR;
              next.posRight = 50 - newWordWidthR;
            } else {
              let dposR = (deltaPercent * 100) / (100 - cw.leftSplit);
              let newWordWidthR = cw.wordRight - dposR;
              newWordWidthR = Math.max(5, Math.min(45, newWordWidthR));
              let newReviewRatioR = (cw.wordRight * cw.reviewRatioRight) / newWordWidthR;

              next.wordRight = newWordWidthR;
              next.posRight = 50 - newWordWidthR;
              next.reviewRatioRight = Math.max(0.1, Math.min(0.9, newReviewRatioR));
            }

            if (isSymmetricSync) {
              next.word = next.wordRight;
              next.pos = next.posRight;
              next.reviewRatio = next.reviewRatioRight;
            }
          }
          return next;
        }

        return next;
      });
    };

    const onMouseUp = () => {
      window.removeEventListener('mousemove', onMouseMove);
      window.removeEventListener('mouseup', onMouseUp);
      setActiveDivider(null);
      document.body.style.cursor = '';
      document.body.style.userSelect = '';
    };

    window.addEventListener('mousemove', onMouseMove);
    window.addEventListener('mouseup', onMouseUp);
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
    storage.setItem('app-enter-newline', String(enterNewLine));
    storage.setItem('app-sync-delete', String(isSyncDelete));
    storage.setItem('app-symmetric-sync', String(isSymmetricSync));
    storage.setItem('app-autosave-interval', String(autoSaveInterval));
    storage.setItem('app-sepia-intensity', String(sepiaIntensity));
    storage.setItem('app-brightness', String(brightness));
    
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
    // Force a small delay to ensure UI reflects current state before export
    await new Promise(resolve => setTimeout(resolve, 100));

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
      const getPara = (text: string, isHeader = false, isCentered?: boolean, useWrap = true) => {
        const processedText = useWrap ? wrapText(text, exportSettings.maxCharsPerLine) : text;
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

      const TABLE_WIDTH_DXA = 9072; // Standard A4 width (11906 total - 1440*2 margin)

      const createCell = (text: string, widthDxa: number, isHeader = false, isCentered?: boolean) => {
        const trimmed = text.trim();
        const isSlash = ['/', '\\', '／'].includes(trimmed);
        
        return new TableCell({
          children: [getPara(isSlash ? '' : text, isHeader, isCentered, false)],
          verticalAlign: VerticalAlign.CENTER,
          width: { size: widthDxa, type: WidthType.DXA },
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
      let mainTableColumnWidths: number[] = [];

      const addSpacer = (size = 200) => {
        sections.push(new Paragraph({
          children: [],
          spacing: { before: size, after: size }
        }));
      };

      if (finalItems.length > 0) {
        addSpacer(100);
        if (exportMode === '4-col') {
          const LS = columnWidths.leftSplit ?? 50;
          const RS = 100 - LS;
          const sumLeft = (columnWidths.word + columnWidths.pos) || 0.1;
          const sumRight = (columnWidths.wordRight + columnWidths.posRight) || 0.1;

          const wLeftWord = Math.round((LS * (columnWidths.word / sumLeft) / 100) * TABLE_WIDTH_DXA);
          const wLeftPos = Math.round((LS * (columnWidths.pos / sumLeft) / 100) * TABLE_WIDTH_DXA);
          const wRightWord = Math.round((RS * (columnWidths.wordRight / sumRight) / 100) * TABLE_WIDTH_DXA);
          const wRightPos = TABLE_WIDTH_DXA - (wLeftWord + wLeftPos + wRightWord); // Ensure sum is exactly TABLE_WIDTH_DXA

          mainTableColumnWidths = [wLeftWord, wLeftPos, wRightWord, wRightPos];

          tableRows.push(new TableRow({
            children: [
              createCell('测试提示', wLeftWord, true), createCell(examType === 'EN' ? '英文' : '中文', wLeftPos, true),
              createCell('测试提示', wRightWord, true), createCell(examType === 'EN' ? '英文' : '中文', wRightPos, true),
            ]
          }));
          for (let i = 0; i < finalItems.length; i += 2) {
            tableRows.push(new TableRow({
              children: [
                createCell(finalItems[i]?.term || '', wLeftWord), createCell('', wLeftPos),
                createCell(finalItems[i+1]?.term || '', wRightWord), createCell('', wRightPos),
              ]
            }));
          }
        } else {
          const LS = columnWidths.leftSplit ?? 50;
          const RS = 100 - LS;
          const sumLeft = (columnWidths.word + columnWidths.pos) || 0.1;
          const sumRight = (columnWidths.wordRight + columnWidths.posRight) || 0.1;

          const width1 = Math.round((LS * (columnWidths.word / sumLeft) * columnWidths.reviewRatio / 100) * TABLE_WIDTH_DXA);
          const width2 = Math.round((LS * (columnWidths.pos / sumLeft) / 100) * TABLE_WIDTH_DXA);
          const width3 = Math.round((LS * (columnWidths.word / sumLeft) * (1 - columnWidths.reviewRatio) / 100) * TABLE_WIDTH_DXA);
          const width4 = Math.round((RS * (columnWidths.wordRight / sumRight) * columnWidths.reviewRatioRight / 100) * TABLE_WIDTH_DXA);
          const width5 = Math.round((RS * (columnWidths.posRight / sumRight) / 100) * TABLE_WIDTH_DXA);
          const width6 = TABLE_WIDTH_DXA - (width1 + width2 + width3 + width4 + width5);

          mainTableColumnWidths = [width1, width2, width3, width4, width5, width6];

          const headerTitle = examType === 'EN' ? '英文' : '中文';
          tableRows.push(new TableRow({
            children: [
              createCell('测试提示', width1, true), createCell('词性', width2, true), createCell(headerTitle, width3, true),
              createCell('测试提示', width4, true), createCell('词性', width5, true), createCell(headerTitle, width6, true),
            ]
          }));

          for (let i = 0; i < finalItems.length; i += 2) {
            const ans1 = exportMode === '6-col' ? '' : (finalItems[i]?.answer || '');
            const ans2 = exportMode === '6-col' ? '' : (finalItems[i+1]?.answer || '');

            tableRows.push(new TableRow({
              children: [
                createCell(finalItems[i]?.term || '', width1), createCell(finalItems[i]?.pos || '', width2), createCell(ans1, width3),
                createCell(finalItems[i+1]?.term || '', width4), createCell(finalItems[i+1]?.pos || '', width5), createCell(ans2, width6),
              ]
            }));
          }
        }
        sections.push(new Table({ 
          width: { size: TABLE_WIDTH_DXA, type: WidthType.DXA }, 
          columnWidths: mainTableColumnWidths,
          layout: TableLayoutType.FIXED,
          rows: tableRows,
          alignment: AlignmentType.CENTER
        }));
        addSpacer(400); 
      }

      if (isExamMode && finalTrans.length > 0) {
        sections.push(new Paragraph({ 
          children: [new TextRun({ text: '变形', bold: true, size: 24, font: 'Microsoft YaHei' })],
          spacing: { before: 200, after: 100 }
        }));
        const transTableRows: TableRow[] = [];
        const transTableWidths = [TABLE_WIDTH_DXA / 2, TABLE_WIDTH_DXA / 2];
        for (let i = 0; i < finalTrans.length; i += 2) {
          const left = finalTrans[i]?.content, right = finalTrans[i+1]?.content;
          transTableRows.push(new TableRow({
            children: [
              createCell(left ? left + "____________________" : "", transTableWidths[0], false, false),
              createCell(right ? right + "____________________" : "", transTableWidths[1], false, false),
            ]
          }));
        }
        sections.push(new Table({ 
          width: { size: TABLE_WIDTH_DXA, type: WidthType.DXA }, 
          columnWidths: transTableWidths,
          layout: TableLayoutType.FIXED,
          rows: transTableRows,
          alignment: AlignmentType.CENTER
        }));
        addSpacer(400);
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

      const doc = new Document({
        sections: [{
          properties: {
            page: {
              size: {
                width: 11906, // A4 Width in twips
                height: 16838, // A4 Height in twips
              },
              margin: {
                top: 1440, // 1 inch
                right: 1440,
                bottom: 1440,
                left: 1440,
              },
            },
          },
          children: sections
        }]
      });
      const blob = await Packer.toBlob(doc);
      saveAs(blob, `${getTodayDate()}.docx`);
    } catch (error) {
      console.error('Export failed:', error);
    } finally {
      setIsExporting(false);
    }
  };

  const handlePOSChange = (e: React.ChangeEvent<HTMLTextAreaElement>): string => {
    let value = e.target.value;
    const selectionStart = e.target.selectionStart;
    
    // 强制转换常见的词性分隔符
    value = value.replace(/[／\\|]/g, '/');
    
    // 自动打点逻辑：仅在增加字符（空格或回车）且光标处于对应位置时触发
    if (value.length > posInput.length && selectionStart > 0) {
      const lastChar = value[selectionStart - 1];

      if (lastChar === ' ' || lastChar === '\n') {
        const { lineIndex } = getLineInfo(value, selectionStart);
        const lines = value.split('\n');
        // 如果是回车触发，目标是上一行；如果是空格触发，目标是当前行
        const targetLineIndex = lastChar === '\n' ? lineIndex - 1 : lineIndex;
        
        if (targetLineIndex >= 0 && targetLineIndex < lines.length) {
          const targetLine = lines[targetLineIndex];
          // 匹配行尾的字母串（忽略末尾可能的空白）
          const match = targetLine.match(/([a-zA-Z]+)(\s*)$/);
          
          if (match && POS_LIST.includes(match[1].toLowerCase())) {
            const word = match[1].toLowerCase();
            const linesBefore = lines.slice(0, targetLineIndex);
            const lineStartPos = linesBefore.join('\n').length + (targetLineIndex > 0 ? 1 : 0);
            const wordInLinePos = targetLine.lastIndexOf(match[1]);
            const absoluteWordPos = lineStartPos + wordInLinePos;
            
            const base = value.slice(0, absoluteWordPos);
            const suffix = lastChar === ' ' ? '. ' : '.\n';
            // 修正 rest 截取点，确保不会丢失其余内容
            const rest = value.slice(selectionStart);
            
            const newVal = base + word + suffix + rest;
            
            // 恢复光标位置
            const newOffset = absoluteWordPos + word.length + suffix.length;
            setTimeout(() => {
              if (posInputRef.current) {
                posInputRef.current.setSelectionRange(newOffset, newOffset);
              }
            }, 0);
            return newVal;
          }
        }
      }
    }
    return value;
  };

  const handleScroll = (e: React.UIEvent<HTMLTextAreaElement>) => {
    const top = e.currentTarget.scrollTop;
    if (scrollRef.current) scrollRef.current.scrollTop = top;
    if (wordInputRef.current && wordInputRef.current.scrollTop !== top) wordInputRef.current.scrollTop = top;
    if (posInputRef.current && posInputRef.current.scrollTop !== top) posInputRef.current.scrollTop = top;
    if (ansInputRef.current && ansInputRef.current.scrollTop !== top) ansInputRef.current.scrollTop = top;
  };

  const downloadOfflineVersion = async () => {
    try {
      // 1. 如果已经在离线版中，基于当前 DOM 生成下载
      if ((window as any).__OFFLINE_VERSION__) {
        const docHtml = document.documentElement.outerHTML;
        const blob = new Blob([docHtml], { type: 'text/html' });
        const url = window.URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = url;
        link.download = 'Randomizer.html';
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        window.URL.revokeObjectURL(url);
        return;
      }

      // 2. 正常下载逻辑
      const response = await fetch(`/offline.html?t=${Date.now()}`);
      if (!response.ok) throw new Error(`HTTP ${response.status}`);
      
      const blob = await response.blob();
      const url = window.URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.download = 'Randomizer.html';
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      window.URL.revokeObjectURL(url);
    } catch (err) {
      console.error("Download failed:", err);
      // 最后尝试直接打开链接
      const link = document.createElement('a');
      link.href = `/offline.html?t=${Date.now()}`;
      link.target = '_blank';
      link.download = 'Randomizer.html';
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
    }
  };

  const computedLineHeight = Math.max(20, Math.round(inputFontSize * 1.857));

  return (
    <div className={cn(
      "min-h-screen font-sans p-6 md:p-12 transition-all duration-300",
      themeStyles.bg
    )} style={theme === 'sepia' ? { backgroundColor: themeStyles.customBg } : {}}>
      <div className="fixed inset-0 pointer-events-none z-[100] bg-black transition-opacity duration-300" style={{ opacity: 1 - brightness / 100 }} />
      <div className="max-w-6xl mx-auto">
        <header className={cn(
          "flex items-end justify-between mb-4 border-b pb-3 transition-colors",
          themeStyles.border
        )}>
          <div>
            <h1 className="text-3xl font-light tracking-tight mb-1 uppercase">无序打乱器</h1>
            <p className={cn("text-xs font-medium", theme === 'light' ? "text-zinc-800" : theme === 'sepia' ? "text-[#5B4636]/80" : "text-zinc-500")}>
              极简单词无序测试制作软件。Produced by Werther with Gemini 3.1 Pro.
            </p>
          </div>
          <div className="flex flex-col items-end gap-3">
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
            <div className="text-right flex flex-col items-end gap-1.5">
              <button 
                onClick={downloadOfflineVersion}
                className={cn(
                  "text-[10px] uppercase font-black px-3 py-1 rounded-lg border transition-all hover:scale-105 active:scale-95 shadow-md",
                  theme === 'light' ? "text-white border-zinc-500 bg-zinc-500 hover:bg-zinc-400 shadow-zinc-200" : 
                  theme === 'sepia' ? "text-[#F4ECD8] border-[#5B4636] bg-[#5B4636] hover:brightness-110 shadow-[#5B4636]/20" : 
                  "text-zinc-100 border-zinc-700 bg-zinc-800 hover:bg-zinc-700 hover:text-white shadow-black/40"
                )}
              >
                下载离线版 (HTML)
              </button>
            </div>
          </div>
        </header>

        <main className="grid grid-cols-1 lg:grid-cols-[1fr_450px] gap-12 items-stretch">
          <section className="flex flex-col h-full gap-6">
            <div className="flex items-center justify-between shrink-0">
              <div className={cn("flex items-center gap-2 text-xs font-bold uppercase tracking-widest", themeStyles.inputText)}>
                <FileText size={14} /> 录入区域
              </div>
              <div className="flex gap-4">
                  <label 
                    className={cn("relative text-[10px] uppercase font-bold transition-all px-3 py-1.5 rounded-xl border active:scale-95 cursor-pointer flex items-center gap-2 shadow-sm", isDragging ? "border-dashed border-cyan-500 bg-cyan-50/10 text-cyan-500 scale-105" : (theme === 'light' ? "text-zinc-800 border-black bg-white hover:bg-gray-50" : theme === 'sepia' ? "text-[#5B4636] border-[#D2C2A4] bg-[#F4ECD8] hover:bg-[#EBE0C5]" : "text-zinc-300 border-zinc-700 bg-zinc-800 hover:bg-zinc-700 hover:text-white"))}
                    onDragOver={onDragOver}
                    onDragLeave={onDragLeave}
                    onDrop={onDrop}
                  >
                    <Download size={12} className="rotate-180" /> 上传文档
                    <input type="file" accept=".txt,.docx" className="hidden" onChange={handleFileUpload} />
                  </label>
                  <button 
                  onClick={() => {
                    setInput("joint efforts\nbe stuck at home\nfacilitate\ninteraction\nreclaim\ncut down\nlessen");
                    setPosInput("phrase.\nphrase.\nv.\nn.\nv.\nphrase.\nv.");
                    setAnswerInput("共同努力\n困在家里\n促进\n互动\n收回\n减少\n减轻/变少");
                    setTransformationInput("comprehend (变形形容词)\nenhance (变形名词)\ndiverse (变形名词)\nacquire (变形名词)");
                    setSentenceInput("1. 我写信是为了推荐一部电影，我的理由如下。\n2. 正如谚语所说：“赠人玫瑰，手有余香。”\n3. 毫无疑问，科技的发展极大地改变了我们的生活。");
                  }}
                  className={cn("text-[10px] uppercase font-bold transition-all px-3 py-1.5 rounded-xl border active:scale-95 shadow-sm", theme === 'light' ? "text-zinc-800 border-black bg-white hover:bg-gray-50" : theme === 'sepia' ? "text-[#5B4636] border-[#D2C2A4] bg-[#F4ECD8] hover:bg-[#EBE0C5]" : "text-zinc-300 border-zinc-700 bg-zinc-800 hover:bg-zinc-700 hover:text-white")}
                >
                  载入示例
                </button>
                <button 
                  onClick={() => { setInput(''); setPosInput(''); setAnswerInput(''); setTransformationInput(''); setSentenceInput(''); }} 
                  className={cn("transition-all p-1.5 rounded-xl border active:scale-95", theme === 'light' ? "border-black text-zinc-400 hover:text-red-500" : theme === 'sepia' ? "border-[#D2C2A4] text-[#8C7B60] hover:text-red-500" : "border-zinc-800 text-gray-300 hover:text-red-500")}
                >
                  <Trash2 size={16} />
                </button>
              </div>
            </div>

            <div className={cn(
              "relative flex rounded-3xl border overflow-hidden min-h-[500px] flex-1 transition-all",
              themeStyles.inputBg, themeStyles.border, "shadow-xl shadow-black/5"
            )} style={theme === 'sepia' ? { backgroundColor: themeStyles.customBg } : {}}>
              {exportMode === '4-col' ? (
                <div className={cn("flex flex-col flex-1 relative group", themeStyles.bg)}>
                  <div className={cn("px-6 py-3 border-b text-[10px] font-bold uppercase tracking-widest", themeStyles.subBg, themeStyles.border, themeStyles.inputText)}>测试提示</div>
                  <div className="flex-1 flex min-h-0">
                    <div ref={scrollRef} className={cn("w-12 border-r text-right pr-3 select-none overflow-hidden transition-colors pointer-events-none block", themeStyles.subBg, themeStyles.border)}>
                      <div style={{ height: '24px' }} className="shrink-0" />
                      {inputLineArray.map((_, i) => (
                        <div key={i} style={{ height: `${computedLineHeight}px`, lineHeight: `${computedLineHeight}px` }} className={cn("text-[11px] flex items-center justify-end font-sans shrink-0", themeStyles.lineNumColor)}>{i + 1}</div>
                      ))}
                      <div style={{ height: '24px' }} className="shrink-0" />
                    </div>
                    <TextAreaInput
                      fontSize={inputFontSize}
                      textareaRef={wordInputRef}
                      value={input}
                      onChange={(e) => setInput(e.target.value)}
                      onScroll={handleScroll}
                      placeholder="输入清单，每行一个..."
                      themeStyles={themeStyles}
                      className={cn("pretty-scrollbar", themeStyles.inputText)}
                      theme={theme}
                      onKeyDown={(e) => {
                        handleBackspaceJump(e, 1);
                        handleEnterJump(e, 1);
                        handleArrowJump(e, 1);
                      }}
                    />
                  </div>
                </div>
              ) : (
                <div id="input-area-container" className={cn(
                  "flex-1 grid relative", 
                  theme === 'dark' ? "gap-px bg-white/20 border-white/20" : theme === 'sepia' ? "gap-px bg-[#D2C2A4] border-[#D2C2A4]" : "gap-px bg-black border-black",
                  "border-b border-t-0"
                )} style={{
                  gridTemplateColumns: exportMode === 'review' ? `${inputColRatio.word}fr ${inputColRatio.pos}fr ${inputColRatio.ans}fr` : `${inputColRatio.word}fr ${inputColRatio.pos}fr`
                }}>
                  <div 
                    onMouseDown={(e) => { e.preventDefault(); setInputActiveDivider('left'); }}
                    className={cn("absolute top-0 bottom-0 w-4 cursor-col-resize z-[60] flex justify-center group", inputActiveDivider === 'left' ? "bg-blue-500/10" : "hover:bg-blue-500/20")}
                    style={{ left: exportMode === 'review' ? `${inputColRatio.word}%` : `${(inputColRatio.word / (inputColRatio.word + inputColRatio.pos)) * 100}%`, transform: 'translateX(-50%)' }}
                  >
                    <div className={cn("w-px h-full transition-colors", inputActiveDivider === 'left' ? "bg-blue-500" : "bg-transparent group-hover:bg-blue-500/50")} />
                  </div>
                  {exportMode === 'review' && (
                    <div 
                      onMouseDown={(e) => { e.preventDefault(); setInputActiveDivider('right'); }}
                      className={cn("absolute top-0 bottom-0 w-4 cursor-col-resize z-[60] flex justify-center group", inputActiveDivider === 'right' ? "bg-blue-500/10" : "hover:bg-blue-500/20")}
                      style={{ left: `${inputColRatio.word + inputColRatio.pos}%`, transform: 'translateX(-50%)' }}
                    >
                      <div className={cn("w-px h-full transition-colors", inputActiveDivider === 'right' ? "bg-blue-500" : "bg-transparent group-hover:bg-blue-500/50")} />
                    </div>
                  )}
                  <div className={cn("flex flex-col relative group", themeStyles.bg)}>
                    <div className={cn("px-6 py-3 border-b text-[10px] font-bold uppercase tracking-widest", themeStyles.subBg, themeStyles.border, themeStyles.inputText)}>测试提示</div>
                    <div className="flex-1 flex min-h-0">
                      <div ref={scrollRef} className={cn("w-10 border-r text-right pr-2 select-none overflow-hidden transition-colors pointer-events-none block scrollbar-hide", themeStyles.subBg, themeStyles.border)}>
                        <div style={{ height: '24px' }} className="shrink-0" />
                        {Array.from({ length: Math.max(inputLineArray.length, posInput.split('\n').length, answerInput.split('\n').length) }).map((_, i) => (
                          <div key={i} style={{ height: `${computedLineHeight}px`, lineHeight: `${computedLineHeight}px` }} className={cn("text-[10px] flex items-center justify-end font-sans shrink-0", themeStyles.lineNumColor)}>{i + 1}</div>
                        ))}
                        <div style={{ height: '24px' }} className="shrink-0" />
                      </div>
                        <TextAreaInput
                          fontSize={inputFontSize}
                          textareaRef={wordInputRef}
                          value={input}
                          onChange={(e) => {
                            const oldVal = input;
                            const newVal = e.target.value;
                            handleSyncDeletions(wordInputRef, oldVal, newVal, setInput, [
                              {val: posInput, setter: setPosInput},
                              {val: answerInput, setter: setAnswerInput}
                            ]);
                          }}
                          onScroll={handleScroll}
                          placeholder="单词..."
                          themeStyles={themeStyles}
                          className="scrollbar-hide"
                          theme={theme}
                          onKeyDown={(e) => {
                            handleBackspaceJump(e, 1);
                            handleEnterJump(e, 1);
                            handleArrowJump(e, 1);
                          }}
                        />
                    </div>
                  </div>
                    <div className={cn("flex flex-col relative group", themeStyles.bg)}>
                      <div className={cn("px-6 py-3 border-b text-[10px] font-bold uppercase tracking-widest", themeStyles.subBg, themeStyles.border, themeStyles.inputText)}>
                        词性
                      </div>
                      <div className="flex-1 flex min-h-0">
                        <TextAreaInput
                          fontSize={inputFontSize}
                          textareaRef={posInputRef}
                          value={posInput}
                          onChange={(e) => {
                            const oldVal = posInput;
                            const newVal = handlePOSChange(e);
                            handleSyncDeletions(posInputRef, oldVal, newVal, setPosInput, [
                              {val: input, setter: setInput},
                              {val: answerInput, setter: setAnswerInput}
                            ]);
                          }}
                          onScroll={handleScroll}
                          className={cn(exportMode === 'review' ? "scrollbar-hide" : "pretty-scrollbar", themeStyles.inputText)}
                          placeholder="词性..."
                          themeStyles={themeStyles}
                          theme={theme}
                       onKeyDown={(e) => {
                         // Handle POS auto-dot before jumping
                         if (e.key === 'Enter' && !e.shiftKey && enterJump) {
                           const { lineIndex } = getLineInfo(e.currentTarget.value, e.currentTarget.selectionStart);
                           let currentVal = e.currentTarget.value;
                           const linesForDot = currentVal.split('\n');
                           const curLine = linesForDot[lineIndex] || '';
                           const match = curLine.match(/([a-zA-Z]+)$/);
                           if (match && POS_LIST.includes(match[1].toLowerCase())) {
                             linesForDot[lineIndex] = curLine + '.';
                             const newVal = linesForDot.join('\n');
                             setPosInput(newVal);
                           }
                         }
                         handleBackspaceJump(e, 2);
                         handleEnterJump(e, 2);
                         handleArrowJump(e, 2);
                       }}
                    />
                    </div>
                  </div>
                  {exportMode === 'review' && (
                    <div className={cn("flex flex-col relative group", themeStyles.bg)}>
                      <div className={cn("px-6 py-3 border-b text-[10px] font-bold uppercase tracking-widest", themeStyles.subBg, themeStyles.border, themeStyles.inputText)}>
                        答案
                      </div>
                      <div className="flex-1 flex min-h-0">
                        <TextAreaInput
                          fontSize={inputFontSize}
                          textareaRef={ansInputRef}
                        value={answerInput}
                        onChange={(e) => {
                          const oldVal = answerInput;
                          const newVal = e.target.value;
                          handleSyncDeletions(ansInputRef, oldVal, newVal, setAnswerInput, [
                            {val: input, setter: setInput},
                            {val: posInput, setter: setPosInput}
                          ]);
                        }}
                        onScroll={handleScroll}
                        className={cn("pretty-scrollbar", themeStyles.inputText)}
                        placeholder="中文答案..."
                        themeStyles={themeStyles}
                        theme={theme}
                        onKeyDown={(e) => {
                          handleBackspaceJump(e, 3);
                          handleEnterJump(e, 3);
                          handleArrowJump(e, 3);
                        }}
                      />
                      </div>
                    </div>
                  )}
                </div>
              )}
            </div>

            {isExamMode && (
              <div className="space-y-6 animate-in slide-in-from-bottom-4 duration-500 shrink-0">
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
                        className={cn("w-full min-h-[150px] bg-transparent resize-none font-sans focus:outline-none leading-relaxed transition-colors", themeStyles.inputText)}
                        style={{ fontSize: `${inputFontSize}px` }}
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
                          setSentenceInput(e.target.value);
                        }}
                        onKeyDown={(e) => {
                          if (e.key === 'Enter' && !e.shiftKey) {
                            const textarea = e.currentTarget;
                            const { lineIndex } = getLineInfo(textarea.value, textarea.selectionStart);
                            const lines = textarea.value.split('\n');
                            const currentLine = lines[lineIndex];
                            
                            // Detect numbering: "1.", "1 ", "1、"
                            const match = currentLine.match(/^(\d+)[\.\s、]/);
                            if (match) {
                              e.preventDefault();
                              const num = parseInt(match[1]);
                              const nextNum = num + 1;
                              const insertText = `\n${nextNum}. `;
                              const start = textarea.selectionStart;
                              const before = textarea.value.slice(0, start);
                              const after = textarea.value.slice(textarea.selectionEnd);
                              
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
                        className={cn("w-full min-h-[150px] bg-transparent resize-none font-mono focus:outline-none leading-relaxed", theme !== 'dark' ? "text-black" : "text-zinc-300")}
                        style={{ fontSize: `${inputFontSize}px` }}
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
                        ? "bg-zinc-500 text-white border-zinc-500 hover:bg-zinc-400"
                        : theme === 'sepia'
                        ? "bg-[#5B4636] text-[#F4ECD8] border-[#5B4636] hover:bg-[#4A392C]"
                        : "bg-zinc-600 text-zinc-100 border-zinc-500 hover:bg-zinc-500"
                    )}
                  >
                    <Layout size={16} className="inline mr-2" /> 预览及导出
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
                        : (input.trim() ? "bg-zinc-700 text-zinc-100 border-zinc-600 hover:bg-zinc-600" : "bg-zinc-800/50 text-zinc-700 border-zinc-800/50 cursor-not-allowed")
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
                        {exportMode === '4-col' && <div className={cn("w-1 h-1 rounded-full", theme === 'dark' ? "bg-zinc-400" : "bg-black")} />}
                      </button>
                      <button onClick={() => setExportMode('6-col')} className={cn("w-full px-4 py-2.5 rounded-xl text-left text-[10px] font-bold transition-all flex items-center justify-between", exportMode === '6-col' ? (themeStyles.inputBg + " " + themeStyles.inputText + " shadow-sm") : (themeStyles.mutedText + " hover:opacity-80"))}>
                        三栏模式（提示+词性）
                        {exportMode === '6-col' && <div className={cn("w-1 h-1 rounded-full", theme === 'dark' ? "bg-zinc-400" : "bg-black")} />}
                      </button>
                      <button onClick={() => setExportMode('review')} className={cn("w-full px-4 py-2.5 rounded-xl text-left text-[10px] font-bold transition-all flex items-center justify-between", exportMode === 'review' ? (themeStyles.inputBg + " " + themeStyles.inputText + " shadow-sm") : (themeStyles.mutedText + " hover:opacity-80"))}>
                        三栏模式（复习资料）
                        {exportMode === 'review' && <div className={cn("w-1 h-1 rounded-full", theme === 'dark' ? "bg-zinc-400" : "bg-black")} />}
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
                              ? (theme === 'light' ? "bg-black border-black" : theme === 'sepia' ? "bg-[#5B4636] border-[#5B4636]" : "bg-zinc-600 border-zinc-500")
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
                              ? (theme === 'light' ? "bg-black border-black" : theme === 'sepia' ? "bg-[#5B4636] border-[#5B4636]" : "bg-zinc-600 border-zinc-500")
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
                        <span className={cn("text-[9px] font-bold uppercase", themeStyles.mutedText)}>回车新建</span>
                        <button 
                          onClick={() => setEnterNewLine(!enterNewLine)}
                          className={cn(
                            "w-8 h-4 rounded-full relative transition-colors border",
                            enterNewLine 
                              ? (theme === 'light' ? "bg-black border-black" : theme === 'sepia' ? "bg-[#5B4636] border-[#5B4636]" : "bg-zinc-600 border-zinc-500")
                              : (themeStyles.toggleBg + " border-black/10")
                          )}
                        >
                          <div className={cn(
                            "absolute top-0.5 w-2.5 h-2.5 rounded-full transition-all",
                            enterNewLine ? "left-4.5 bg-white" : "left-0.5 bg-gray-500"
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
                              ? (theme === 'light' ? "bg-black border-black" : theme === 'sepia' ? "bg-[#5B4636] border-[#5B4636]" : "bg-zinc-600 border-zinc-500")
                              : (themeStyles.toggleBg + " border-black/10")
                          )}
                        >
                          <div className={cn(
                            "absolute top-0.5 w-2.5 h-2.5 rounded-full transition-all",
                            isSyncDelete ? "left-4.5 bg-white" : "left-0.5 bg-gray-500"
                          )} />
                        </button>
                      </div>
                      
                      <div className="flex items-center justify-between gap-4">
                        <span className={cn("text-[9px] font-bold uppercase", themeStyles.mutedText)}>录入区布局</span>
                        <button
                          onClick={() => {
                            if (exportMode === '4-col') {
                              setInputColRatio({ word: 50, pos: 50, ans: 0 });
                            } else {
                              setInputColRatio({ word: 33.33, pos: 33.33, ans: 33.34 });
                            }
                          }}
                          className={cn("px-2 py-1 text-[9px] font-bold rounded-md transition-all border", themeStyles.inputBg, themeStyles.inputText, "hover:opacity-80")}
                        >
                          恢复默认
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
                      <div className="flex items-center justify-between">
                        <span className={cn("text-[9px] font-bold uppercase", themeStyles.mutedText)}>自动保存 (秒)</span>
                        <div className="flex items-center gap-2">
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
                      
                      <div className="flex items-center justify-between">
                        <span className={cn("text-[9px] font-bold uppercase", themeStyles.mutedText)}>视觉大小</span>
                        <div className="flex items-center gap-2">
                          <input 
                            type="range"
                            min="12"
                            max="24"
                            step="1"
                            value={inputFontSize}
                            onChange={(e) => setInputFontSize(Number(e.target.value))}
                            className={cn("w-20 h-1 rounded-full appearance-none cursor-pointer", theme === 'dark' ? "bg-white/10 accent-white" : theme === 'sepia' ? "bg-[#5B4636]/20 accent-[#5B4636]" : "bg-black/10 accent-black")}
                          />
                          <span className={cn("text-[10px] font-bold min-w-[24px] text-right", themeStyles.inputText)}>{inputFontSize}px</span>
                        </div>
                      </div>

                      <div className="flex items-center justify-between">
                        <span className={cn("text-[9px] font-bold uppercase", themeStyles.mutedText)}>调节亮度</span>
                        <div className="flex items-center gap-2">
                          <input 
                            type="range" 
                            min="30" 
                            max="100" 
                            value={brightness} 
                            onChange={(e) => setBrightness(Number(e.target.value))}
                            className={cn("w-20 h-1 rounded-full appearance-none cursor-pointer", theme === 'dark' ? "bg-white/10 accent-white" : theme === 'sepia' ? "bg-[#5B4636]/20 accent-[#5B4636]" : "bg-black/10 accent-black")}
                          />
                          <span className={cn("text-[10px] font-bold min-w-[24px] text-right", themeStyles.inputText)}>{brightness}%</span>
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
                    ? "bg-zinc-500 text-white border-zinc-500 hover:bg-zinc-400"
                    : theme === 'sepia'
                    ? "bg-[#5B4636] text-[#F4ECD8] border-[#5B4636] hover:bg-[#4A392C]"
                    : "bg-zinc-800 text-zinc-100 border-zinc-700 hover:bg-zinc-700 hover:text-white"
                )}
              >
                {isSavingDefaults ? "已保存默认配置 ✔" : "保存当前偏好为默认"}
              </button>
            </div>
          </aside>
        </main>

        <AnimatePresence>
          {showStylePicker && (
            <motion.div 
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              className="fixed inset-0 z-[200] flex items-center justify-center p-6 bg-black/40 backdrop-blur-md"
            >
              <motion.div 
                initial={{ scale: 0.9, opacity: 0, y: 20 }}
                animate={{ scale: 1, opacity: 1, y: 0 }}
                exit={{ scale: 0.9, opacity: 0, y: 20 }}
                className={cn("max-w-2xl w-full p-10 rounded-[40px] border shadow-2xl space-y-8 relative overflow-hidden", themeStyles.inputBg, themeStyles.border)}
                style={theme === 'sepia' ? { backgroundColor: themeStyles.customBg } : {}}
              >
                <div className="absolute top-0 right-0 p-8">
                  <button onClick={() => setShowStylePicker(false)} className={cn("opacity-40 hover:opacity-100 transition-opacity", themeStyles.inputText)}>
                     <X size={20} />
                  </button>
                </div>
                
                <div className="text-center space-y-2">
                  <h2 className={cn("text-3xl font-light tracking-tight", themeStyles.inputText)}>选择预览样式</h2>
                  <p className={cn("text-sm transition-opacity", theme === 'dark' ? "text-zinc-500" : "text-gray-500")}>请选择导出内容的呈现方案</p>
                </div>

                <div className="grid grid-cols-1 md:grid-cols-2 gap-8">
                  <button 
                    onClick={() => startPreview('table')}
                    className={cn(
                      "group relative overflow-hidden rounded-[32px] border p-8 text-left transition-all hover:scale-[1.02] active:scale-[0.98] shadow-sm",
                      themeStyles.subBg, themeStyles.border, theme === 'light' ? "hover:border-black" : "hover:border-white/40"
                    )}
                  >
                    <div className="aspect-video mb-6 rounded-2xl border flex items-center justify-center overflow-hidden relative bg-zinc-400/5 group-hover:bg-zinc-400/10 transition-colors">
                      {/* Mini Table Preview */}
                      <div className="w-full h-full p-4 flex flex-col gap-2">
                        <div className="flex gap-2 opacity-40">
                          <div className="h-4 w-1/3 bg-current rounded-md" />
                          <div className="h-4 w-1/6 bg-current rounded-md" />
                          <div className="h-4 w-1/2 bg-current rounded-md" />
                        </div>
                        <div className="h-[1px] w-full bg-current opacity-20" />
                        {[1, 2, 3].map(i => (
                          <div key={i} className="flex gap-2 opacity-20">
                            <div className="h-3 w-1/3 border rounded-sm" />
                            <div className="h-3 w-1/6 border rounded-sm" />
                            <div className="h-3 w-1/2 border rounded-sm" />
                          </div>
                        ))}
                      </div>
                    </div>
                    <div className="flex items-center justify-between mb-2">
                      <h3 className={cn("font-bold text-lg", themeStyles.inputText)}>精美表格模式</h3>
                      <div className={cn("p-1.5 rounded-full", themeStyles.subBg)}>
                         <Layout size={16} className={themeStyles.inputText} />
                      </div>
                    </div>
                    <p className="text-[11px] text-zinc-500 leading-relaxed font-medium">包含完整列表、词性标注与多列排版，支持 Word 格式下载。适合正式测试打印使用。</p>
                  </button>

                  <button 
                    onClick={() => startPreview('text')}
                    className={cn(
                      "group relative overflow-hidden rounded-[32px] border p-8 text-left transition-all hover:scale-[1.02] active:scale-[0.98] shadow-sm",
                      themeStyles.subBg, themeStyles.border, theme === 'light' ? "hover:border-black" : "hover:border-white/40"
                    )}
                  >
                    <div className="aspect-video mb-6 rounded-2xl border flex items-center justify-center overflow-hidden relative bg-zinc-400/5 group-hover:bg-zinc-400/10 transition-colors">
                      {/* Mini Text Preview */}
                      <div className="w-full h-full p-4 flex flex-col gap-2 font-mono">
                        {[1, 2, 3, 4].map(i => (
                          <div key={i} className="flex gap-2 opacity-30">
                            <div className="h-2 w-1/4 bg-current rounded-full" />
                            <div className="h-2 w-3 bg-current rounded-full" />
                            <div className="h-2 w-1/2 bg-current rounded-full" />
                          </div>
                        ))}
                      </div>
                    </div>
                    <div className="flex items-center justify-between mb-2">
                      <h3 className={cn("font-bold text-lg", themeStyles.inputText)}>简约文本模式</h3>
                      <div className={cn("p-1.5 rounded-full", themeStyles.subBg)}>
                         <Type size={16} className={themeStyles.inputText} />
                      </div>
                    </div>
                    <p className="text-[11px] text-zinc-500 leading-relaxed font-medium">每行呈现「提示 词性 答案」，适合快速复制。直接生成文本块，方便粘贴到其他文档。</p>
                  </button>
                </div>
              </motion.div>
            </motion.div>
          )}
        </AnimatePresence>

        {showPreview && (
          <div className="fixed inset-0 z-50 flex flex-col bg-black/80 backdrop-blur-md animate-in fade-in duration-300">
            <div className={cn(
              "flex-1 flex flex-col w-full h-full border-b shadow-2xl overflow-hidden",
              theme === 'dark' ? "bg-[#111112] border-zinc-800 text-[#E4E4E7]" : 
              theme === 'sepia' ? "border-[#D2C2A4] text-[#5B4636]" : 
              "bg-[#F5F5F7] border-black"
            )}
            style={theme === 'sepia' ? { backgroundColor: themeStyles.customBg } : {}}>
              {/* Header */}
              <div className={cn("p-6 md:px-12 md:py-8 border-b flex justify-between items-center", 
                theme === 'dark' ? "border-zinc-800" : 
                theme === 'sepia' ? "border-[#D2C2A4]/50" : 
                "border-black/10"
              )}>
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
                      theme === 'dark' ? "bg-zinc-600 text-white border-zinc-500 hover:bg-zinc-500 font-bold" : 
                      theme === 'sepia' ? "bg-[#5B4636] text-[#F4ECD8] border-[#5B4636] hover:bg-[#4A392C]" :
                      "bg-zinc-500 text-white border-zinc-500 hover:bg-zinc-400"
                    )}
                  >
                    {isExporting ? <div className="w-4 h-4 border-2 border-zinc-200 border-t-zinc-800 rounded-full animate-spin" /> : <Download size={14} />} 导出 WORD
                  </button>
                  <button 
                    onClick={() => setShowPreview(false)}
                    className={cn(
                      "p-3 rounded-2xl transition-all border shadow-sm",
                      theme === 'dark' ? "bg-zinc-800 border-zinc-700 hover:bg-zinc-700 text-white" : 
                      theme === 'sepia' ? "bg-[#EBE0C5] border-[#D2C2A4] hover:bg-[#D2C2A4] text-[#5B4636]" :
                      "bg-white border-zinc-300 hover:bg-zinc-50 text-black"
                    )}
                  >
                    <X size={20} />
                  </button>
                </div>
              </div>

              {/* Body */}
              <div className="flex-1 overflow-y-auto p-6 md:p-12 space-y-12 pretty-scrollbar backdrop-blur-sm">
                {exportStyle === 'text' ? (
                  <div className="max-w-4xl mx-auto space-y-8 animate-in slide-in-from-bottom-4 duration-500">
                    <div className={cn("p-10 rounded-[40px] border shadow-sm", themeStyles.innerBg, themeStyles.borderMuted)} style={theme === 'sepia' ? { backgroundColor: themeStyles.customInnerBg } : {}}>
                      <div className="flex items-center justify-between mb-8">
                        <div className="flex items-center gap-3">
                          <div className={cn("p-3 rounded-2xl", themeStyles.subBg)} style={theme === 'sepia' ? { backgroundColor: themeStyles.customSubBg } : {}}>
                             <Copy size={24} className={themeStyles.inputText} />
                          </div>
                          <div>
                            <h3 className={cn("text-xl font-light", themeStyles.inputText)}>纯文本生成结果</h3>
                            <p className="text-[10px] uppercase font-bold text-zinc-500 tracking-widest mt-1">格式：提示词 + 词性 + 答案</p>
                          </div>
                        </div>
                        <button 
                          onClick={() => {
                            const text = previewItems.map(item => `${item.term}${item.pos ? ` ${item.pos}` : ''}${item.answer ? ` ${item.answer}` : ''}`).join('\n');
                            navigator.clipboard.writeText(text);
                            alert('内容已成功复制到剪贴板！');
                          }}
                          className={cn("px-6 py-2.5 rounded-xl text-[10px] font-black uppercase tracking-widest transition-all active:scale-95", 
                            theme === 'light' ? "bg-black text-white hover:bg-zinc-800" : 
                            theme === 'sepia' ? "bg-[#5B4636] text-[#F4ECD8] hover:bg-[#4A392C]" : 
                            "bg-white text-black hover:bg-zinc-100"
                          )}
                        >
                          一键复制全文
                        </button>
                      </div>
                      <div className={cn("rounded-3xl border p-8 font-mono text-sm leading-loose overflow-x-auto whitespace-pre h-[500px] pretty-scrollbar", themeStyles.subBg, themeStyles.border, themeStyles.inputText)}>
                        {previewItems.map(item => `${item.term}${item.pos ? ` ${item.pos}` : ''}${item.answer ? ` ${item.answer}` : ''}`).join('\n')}
                      </div>
                    </div>
                  </div>
                ) : (
                  <>
                    {/* Spacing Controls Bar */}
                <div className={cn(
                  "p-6 rounded-[32px] border flex flex-wrap items-center gap-8",
                  themeStyles.innerBg, themeStyles.border
                )} style={theme === 'sepia' ? { backgroundColor: themeStyles.customInnerBg } : {}}>
                  <div className="flex items-center gap-4">
                    <span className={cn("text-[10px] font-bold uppercase flex items-center gap-2", themeStyles.mutedText)}>
                       <Layout size={12} /> 行间距
                    </span>
                    <div className={cn("flex items-center border rounded-xl overflow-hidden px-3 py-1", themeStyles.innerBg, themeStyles.border)} style={theme === 'sepia' ? { backgroundColor: themeStyles.customInnerBg } : {}}>
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
                    <div className={cn("flex items-center border rounded-xl overflow-hidden px-3 py-1", themeStyles.innerBg, themeStyles.border)} style={theme === 'sepia' ? { backgroundColor: themeStyles.customInnerBg } : {}}>
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

                  <div className="flex items-center gap-4">
                    <span className={cn("text-[10px] font-bold uppercase flex items-center gap-2", themeStyles.mutedText)}>
                       考察内容
                    </span>
                    <div className={cn("flex p-0.5 rounded-lg border", themeStyles.innerBg, themeStyles.border)} style={theme === 'sepia' ? { backgroundColor: themeStyles.customInnerBg } : {}}>
                      <button 
                        onClick={() => setExamType('EN')}
                        className={cn("px-2 py-1 rounded-md text-[9px] font-bold transition-all", examType === 'EN' ? theme === 'dark' ? "bg-white text-black shadow-sm" : theme === 'sepia' ? "bg-[#5B4636] text-[#F4ECD8] shadow-sm" : "bg-black text-white shadow-sm" : "opacity-50")}
                      >英文</button>
                      <button 
                        onClick={() => setExamType('CN')}
                        className={cn("px-2 py-1 rounded-md text-[9px] font-bold transition-all", examType === 'CN' ? theme === 'dark' ? "bg-white text-black shadow-sm" : theme === 'sepia' ? "bg-[#5B4636] text-[#F4ECD8] shadow-sm" : "bg-black text-white shadow-sm" : "opacity-50")}
                      >中文</button>
                    </div>
                  </div>

                  <div className="flex items-center gap-4">
                    <span className={cn("text-[10px] font-bold uppercase flex items-center gap-2", themeStyles.mutedText)}>
                       对称调整列宽
                    </span>
                    <button 
                      onClick={() => setIsSymmetricSync(!isSymmetricSync)}
                      className={cn(
                        "w-8 h-4 rounded-full relative transition-colors border",
                        isSymmetricSync 
                          ? (theme === 'light' ? "bg-black border-black" : theme === 'sepia' ? "bg-[#5B4636] border-[#5B4636]" : "bg-zinc-600 border-zinc-500")
                          : (themeStyles.innerBg + " " + themeStyles.border)
                      )}
                      style={!isSymmetricSync && theme === 'sepia' ? { backgroundColor: themeStyles.customInnerBg } : {}}
                    >
                      <div className={cn(
                        "absolute top-0.5 w-2.5 h-2.5 rounded-full transition-all",
                        isSymmetricSync ? "left-4.5 bg-white" : "left-0.5 bg-gray-500"
                      )} />
                    </button>
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
                    <SortableContext items={previewItems} strategy={rectSortingStrategy}>
                      <div 
                        ref={previewTableRef}
                        className={cn(
                          "flex flex-wrap border-l border-t overflow-hidden rounded-xl relative",
                          theme === 'dark' ? "border-white/20 bg-black" : theme === 'sepia' ? "border-[#D2C2A4] bg-transparent" : "border-black bg-transparent"
                        )}
                        style={{ width: '100%' }}
                      >
                        {/* Header Row */}
                        <div className="w-full flex">
                          {/* Left Headers */}
                          <div style={{ width: `${columnWidths.leftSplit}%` }} className="flex">
                            {exportMode === '4-col' ? (
                              <>
                                <div className={cn("p-3 text-[10px] font-bold text-center border-r border-b border-current uppercase", themeStyles.subBg)} style={theme === 'sepia' ? { backgroundColor: themeStyles.customSubBg, flex: columnWidths.word } : { flex: columnWidths.word }}>测试提示</div>
                                <div className={cn("p-3 text-[10px] font-bold text-center border-r border-b border-current uppercase", themeStyles.subBg)} style={theme === 'sepia' ? { backgroundColor: themeStyles.customSubBg, flex: columnWidths.pos || 0.1 } : { flex: columnWidths.pos || 0.1 }}>{examType === 'EN' ? '英文' : '中文'}</div>
                              </>
                            ) : (
                              <>
                                <div className={cn("p-3 text-[10px] font-bold text-center border-r border-b border-current uppercase", themeStyles.subBg)} style={theme === 'sepia' ? { backgroundColor: themeStyles.customSubBg, flex: columnWidths.word * columnWidths.reviewRatio } : { flex: columnWidths.word * columnWidths.reviewRatio }}>测试提示</div>
                                <div className={cn("p-3 text-[10px] font-bold text-center border-r border-b border-current uppercase", themeStyles.subBg)} style={theme === 'sepia' ? { backgroundColor: themeStyles.customSubBg, flex: columnWidths.pos || 0.1 } : { flex: columnWidths.pos || 0.1 }}>词性</div>
                                <div className={cn("p-3 text-[10px] font-bold text-center border-r border-b border-current uppercase", themeStyles.subBg)} style={theme === 'sepia' ? { backgroundColor: themeStyles.customSubBg, flex: columnWidths.word * (1 - columnWidths.reviewRatio) } : { flex: columnWidths.word * (1 - columnWidths.reviewRatio) }}>{examType === 'EN' ? '英文' : '中文'}</div>
                              </>
                            )}
                          </div>
                          {/* Right Headers */}
                          <div style={{ width: `${100 - columnWidths.leftSplit}%` }} className="flex">
                            {exportMode === '4-col' ? (
                              <>
                                <div className={cn("p-3 text-[10px] font-bold text-center border-r border-b border-current uppercase", themeStyles.subBg)} style={theme === 'sepia' ? { backgroundColor: themeStyles.customSubBg, flex: columnWidths.wordRight } : { flex: columnWidths.wordRight }}>测试提示</div>
                                <div className={cn("p-3 text-[10px] font-bold text-center border-r border-b border-current uppercase", themeStyles.subBg)} style={theme === 'sepia' ? { backgroundColor: themeStyles.customSubBg, flex: columnWidths.posRight || 0.1 } : { flex: columnWidths.posRight || 0.1 }}>{examType === 'EN' ? '英文' : '中文'}</div>
                              </>
                            ) : (
                              <>
                                <div className={cn("p-3 text-[10px] font-bold text-center border-r border-b border-current uppercase", themeStyles.subBg)} style={theme === 'sepia' ? { backgroundColor: themeStyles.customSubBg, flex: columnWidths.wordRight * columnWidths.reviewRatioRight } : { flex: columnWidths.wordRight * columnWidths.reviewRatioRight }}>测试提示</div>
                                <div className={cn("p-3 text-[10px] font-bold text-center border-r border-b border-current uppercase", themeStyles.subBg)} style={theme === 'sepia' ? { backgroundColor: themeStyles.customSubBg, flex: columnWidths.posRight || 0.1 } : { flex: columnWidths.posRight || 0.1 }}>词性</div>
                                <div className={cn("p-3 text-[10px] font-bold text-center border-r border-b border-current uppercase", themeStyles.inputText, themeStyles.subBg)} style={theme === 'sepia' ? { backgroundColor: themeStyles.customSubBg, flex: columnWidths.wordRight * (1 - columnWidths.reviewRatioRight) } : { flex: columnWidths.wordRight * (1 - columnWidths.reviewRatioRight) }}>{examType === 'EN' ? '英文' : '中文'}</div>
                              </>
                            )}
                          </div>
                        </div>

                        {/* Data Rows - Render flat array of previewItems */}
                        {previewItems.map((item, index) => {
                          const isLeft = index % 2 === 0;
                          const widthPct = isLeft ? columnWidths.leftSplit : (100 - columnWidths.leftSplit);
                          
                          const wordRatio = isLeft ? columnWidths.word : columnWidths.wordRight;
                          const posRatio = isLeft ? (columnWidths.pos || 0.1) : (columnWidths.posRight || 0.1);
                          const revRatio = isLeft ? columnWidths.reviewRatio : columnWidths.reviewRatioRight;

                          return (
                            <SortableItem 
                              key={item.id} 
                              id={item.id} 
                              className="flex transition-colors w-full" 
                              style={{ width: `${widthPct}%` }}
                            >
                              <div className={cn("p-3 text-xs border-r border-b font-medium text-center break-all min-h-[40px] flex items-center justify-center", themeStyles.borderMuted, themeStyles.subBg)} style={theme === 'sepia' ? { backgroundColor: themeStyles.customSubBg, flex: exportMode === '4-col' ? wordRatio : wordRatio * revRatio } : { flex: exportMode === '4-col' ? wordRatio : wordRatio * revRatio }}>
                                <span className="opacity-80 font-serif">{item.term}</span>
                              </div>
                              
                              {exportMode === '6-col' || exportMode === 'review' ? (
                                <div className={cn("p-3 text-[10px] font-mono border-b border-r italic text-center text-zinc-500 break-all min-h-[40px] flex items-center justify-center", themeStyles.borderMuted, themeStyles.subBg)} style={theme === 'sepia' ? { backgroundColor: themeStyles.customSubBg, flex: posRatio } : { flex: posRatio }}>
                                  {item.pos}
                                </div>
                              ) : null}

                              {exportMode === 'review' && (
                                <div className={cn("p-3 text-xs text-center break-all border-b border-r min-h-[40px] flex items-center justify-center", themeStyles.borderMuted, theme === 'sepia' ? "" : themeStyles.bg)} style={theme === 'sepia' ? { backgroundColor: themeStyles.customBg, flex: wordRatio * (1 - revRatio) } : { flex: wordRatio * (1 - revRatio) }}>
                                  {item.answer}
                                </div>
                              )}
                              
                              {exportMode === '6-col' && (
                                <div className={cn("p-3 text-xs text-center break-all border-b border-r min-h-[40px]", themeStyles.borderMuted, theme === 'sepia' ? "" : themeStyles.bg)} style={theme === 'sepia' ? { backgroundColor: themeStyles.customBg, flex: wordRatio * (1 - revRatio) } : { flex: wordRatio * (1 - revRatio) }} />
                              )}
                              
                              {exportMode === '4-col' && (
                                <div className={cn("p-3 text-xs text-center break-all border-b border-r min-h-[40px] flex items-center justify-center", themeStyles.borderMuted, theme === 'sepia' ? "" : themeStyles.bg)} style={theme === 'sepia' ? { backgroundColor: themeStyles.customBg, flex: posRatio } : { flex: posRatio }} />
                              )}
                            </SortableItem>
                          );
                        })}
                        
                        {/* Filler cells to complete the row if odd items! */}
                        {previewItems.length % 2 !== 0 && (() => {
                          const isLeft = false; // The filler is always on the right if it's odd items!
                          const widthPct = 100 - columnWidths.leftSplit;
                          const wordRatio = columnWidths.wordRight;
                          const posRatio = columnWidths.posRight || 0.1;
                          const revRatio = columnWidths.reviewRatioRight;

                          return (
                            <div className="flex transition-colors w-full" style={{ width: `${widthPct}%` }}>
                              <div className={cn("p-3 text-xs border-r border-b min-h-[40px]", themeStyles.borderMuted, themeStyles.subBg)} style={theme === 'sepia' ? { backgroundColor: themeStyles.customSubBg, flex: exportMode === '4-col' ? wordRatio : wordRatio * revRatio } : { flex: exportMode === '4-col' ? wordRatio : wordRatio * revRatio }} />
                              
                              {exportMode === '6-col' || exportMode === 'review' ? (
                                <div className={cn("p-3 text-[10px] border-b border-r min-h-[40px]", themeStyles.borderMuted, themeStyles.subBg)} style={theme === 'sepia' ? { backgroundColor: themeStyles.customSubBg, flex: posRatio } : { flex: posRatio }} />
                              ) : null}

                              {exportMode === 'review' && (
                                <div className={cn("p-3 text-xs border-b border-r min-h-[40px]", themeStyles.borderMuted, theme === 'sepia' ? "" : themeStyles.bg)} style={theme === 'sepia' ? { backgroundColor: themeStyles.customBg, flex: wordRatio * (1 - revRatio) } : { flex: wordRatio * (1 - revRatio) }} />
                              )}
                              
                              {exportMode === '6-col' && (
                                <div className={cn("p-3 text-xs border-b border-r min-h-[40px]", themeStyles.borderMuted, theme === 'sepia' ? "" : themeStyles.bg)} style={theme === 'sepia' ? { backgroundColor: themeStyles.customBg, flex: wordRatio * (1 - revRatio) } : { flex: wordRatio * (1 - revRatio) }} />
                              )}
                              
                              {exportMode === '4-col' && (
                                <div className={cn("p-3 text-xs border-b border-r min-h-[40px]", themeStyles.borderMuted, theme === 'sepia' ? "" : themeStyles.bg)} style={theme === 'sepia' ? { backgroundColor: themeStyles.customBg, flex: posRatio } : { flex: posRatio }} />
                              )}
                            </div>
                          );
                        })()}
                        
                        {/* Drag Handle Overlays - Moved after content for best layering */}
                        {exportMode === '4-col' && (
                          <>
                            <div 
                              onMouseDown={(e) => handleDividerDrag(e, 'main', 'left')}
                              className={cn("absolute top-0 bottom-0 w-4 cursor-col-resize z-[60] flex justify-center group", activeDivider?.type === 'main' && activeDivider?.side === 'left' ? "bg-blue-500/10" : "hover:bg-blue-500/20")}
                              style={{ left: `${(columnWidths.leftSplit ?? 50) * (columnWidths.word / 50)}%`, transform: 'translateX(-50%)' }}
                            >
                              <div className={cn("w-px h-full transition-colors", activeDivider?.type === 'main' && activeDivider?.side === 'left' ? "bg-blue-500" : "bg-transparent group-hover:bg-blue-500/50")} />
                            </div>
                            <div 
                              onMouseDown={(e) => handleDividerDrag(e, 'split', 'center')}
                              className={cn("absolute top-0 bottom-0 w-4 cursor-col-resize z-[60] flex justify-center group", activeDivider?.type === 'split' ? "bg-blue-500/10" : "hover:bg-blue-500/20")}
                              style={{ left: `${columnWidths.leftSplit ?? 50}%`, transform: 'translateX(-50%)' }}
                            >
                              <div className={cn("w-px h-full transition-colors", activeDivider?.type === 'split' ? "bg-blue-500" : "bg-transparent group-hover:bg-blue-500/50")} />
                            </div>
                            <div 
                              onMouseDown={(e) => handleDividerDrag(e, 'main', 'right')}
                              className={cn("absolute top-0 bottom-0 w-4 cursor-col-resize z-[60] flex justify-center group", activeDivider?.type === 'main' && activeDivider?.side === 'right' ? "bg-blue-500/10" : "hover:bg-blue-500/20")}
                              style={{ left: `${(columnWidths.leftSplit ?? 50) + (100 - (columnWidths.leftSplit ?? 50)) * (columnWidths.wordRight / 50)}%`, transform: 'translateX(-50%)' }}
                            >
                              <div className={cn("w-px h-full transition-colors", activeDivider?.type === 'main' && activeDivider?.side === 'right' ? "bg-blue-500" : "bg-transparent group-hover:bg-blue-500/50")} />
                            </div>
                          </>
                        )}
                        {exportMode !== '4-col' && (
                          <>
                            <div 
                              onMouseDown={(e) => handleDividerDrag(e, 'review', 'left')}
                              className={cn("absolute top-0 bottom-0 w-4 cursor-col-resize z-[60] flex justify-center group", activeDivider?.type === 'review' && activeDivider?.side === 'left' ? "bg-blue-500/10" : "hover:bg-blue-500/20")}
                              style={{ left: `${(columnWidths.leftSplit ?? 50) * (columnWidths.word / 50) * columnWidths.reviewRatio}%`, transform: 'translateX(-50%)' }}
                            >
                              <div className={cn("w-px h-full transition-colors", activeDivider?.type === 'review' && activeDivider?.side === 'left' ? "bg-blue-500" : "bg-transparent group-hover:bg-blue-500/50")} />
                            </div>
                            <div 
                              onMouseDown={(e) => handleDividerDrag(e, 'main', 'left')}
                              className={cn("absolute top-0 bottom-0 w-4 cursor-col-resize z-[60] flex justify-center group", activeDivider?.type === 'main' && activeDivider?.side === 'left' ? "bg-blue-500/10" : "hover:bg-blue-500/20")}
                              style={{ left: `${(columnWidths.leftSplit ?? 50) * ((columnWidths.word / 50) * columnWidths.reviewRatio + (columnWidths.pos / 50))}%`, transform: 'translateX(-50%)' }}
                            >
                              <div className={cn("w-px h-full transition-colors", activeDivider?.type === 'main' && activeDivider?.side === 'left' ? "bg-blue-500" : "bg-transparent group-hover:bg-blue-500/50")} />
                            </div>
                            <div 
                              onMouseDown={(e) => handleDividerDrag(e, 'split', 'center')}
                              className={cn("absolute top-0 bottom-0 w-4 cursor-col-resize z-[60] flex justify-center group", activeDivider?.type === 'split' ? "bg-blue-500/10" : "hover:bg-blue-500/20")}
                              style={{ left: `${columnWidths.leftSplit ?? 50}%`, transform: 'translateX(-50%)' }}
                            >
                              <div className={cn("w-px h-full transition-colors", activeDivider?.type === 'split' ? "bg-blue-500" : "bg-transparent group-hover:bg-blue-500/50")} />
                            </div>
                            <div 
                              onMouseDown={(e) => handleDividerDrag(e, 'review', 'right')}
                              className={cn("absolute top-0 bottom-0 w-4 cursor-col-resize z-[60] flex justify-center group", activeDivider?.type === 'review' && activeDivider?.side === 'right' ? "bg-blue-500/10" : "hover:bg-blue-500/20")}
                              style={{ left: `${(columnWidths.leftSplit ?? 50) + (100 - (columnWidths.leftSplit ?? 50)) * (columnWidths.wordRight / 50) * columnWidths.reviewRatioRight}%`, transform: 'translateX(-50%)' }}
                            >
                              <div className={cn("w-px h-full transition-colors", activeDivider?.type === 'review' && activeDivider?.side === 'right' ? "bg-blue-500" : "bg-transparent group-hover:bg-blue-500/50")} />
                            </div>
                            <div 
                              onMouseDown={(e) => handleDividerDrag(e, 'main', 'right')}
                              className={cn("absolute top-0 bottom-0 w-4 cursor-col-resize z-[60] flex justify-center group", activeDivider?.type === 'main' && activeDivider?.side === 'right' ? "bg-blue-500/10" : "hover:bg-blue-500/20")}
                              style={{ left: `${(columnWidths.leftSplit ?? 50) + (100 - (columnWidths.leftSplit ?? 50)) * ((columnWidths.wordRight / 50) * columnWidths.reviewRatioRight + (columnWidths.posRight / 50))}%`, transform: 'translateX(-50%)' }}
                            >
                              <div className={cn("w-px h-full transition-colors", activeDivider?.type === 'main' && activeDivider?.side === 'right' ? "bg-blue-500" : "bg-transparent group-hover:bg-blue-500/50")} />
                            </div>
                          </>
                        )}
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
                          theme === 'dark' ? "border-zinc-700 bg-zinc-800" : theme === 'sepia' ? "border-[#D2C2A4] bg-[#D2C2A4]" : "border-black bg-black"
                        )}>
                          {previewTransformations.map((t) => (
                            <SortableItem key={t.id} id={t.id} className={cn("p-4 text-xs", theme === 'dark' ? "bg-zinc-950" : theme === 'sepia' ? "" : "bg-white")} style={theme === 'sepia' ? { backgroundColor: themeStyles.customBg } : {}}>
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
                            <SortableItem key={s.id} id={s.id} className={cn("p-6 rounded-3xl border group transition-all hover:border-black/40", theme === 'dark' ? "bg-zinc-900/40 border-zinc-800/50" : theme === 'sepia' ? "border-[#D2C2A4]/50" : "bg-white border-black/10")} style={theme === 'sepia' ? { backgroundColor: themeStyles.customBg } : {}}>
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
                  </>
                )}
                
                <div className="h-24"></div>
                {exportStyle === 'text' && <div className="h-48" />}
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
