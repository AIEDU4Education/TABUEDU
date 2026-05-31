import React, { useState, useEffect, useRef } from 'react';
import * as XLSX from 'xlsx';
import confetti from 'canvas-confetti';
import { motion, AnimatePresence } from 'motion/react';
import { 
  Trophy, 
  Award,
  Medal,
  Crown,
  Settings, 
  Play, 
  RotateCcw, 
  Plus, 
  Minus, 
  X,
  Check,
  SkipForward, 
  Upload, 
  Download,
  FileSpreadsheet,
  RefreshCw,
  Trash2,
  Sparkles, 
  Timer, 
  Users, 
  Layers,
  Lock,
  LockOpen,
  ChevronLeft,
  ChevronRight,
  Info,
  CheckCircle2,
  AlertCircle,
  School,
  FileText,
  Eye,
  EyeOff,
  Flag,
  LogOut,
  Volume2,
  Target,
  Pencil
} from 'lucide-react';
import { 
  Document, 
  Packer, 
  Paragraph, 
  Table, 
  TableRow, 
  TableCell, 
  WidthType, 
  AlignmentType, 
  HeadingLevel,
  VerticalAlign,
  TextRun
} from 'docx';
import { saveAs } from 'file-saver';
import { Button } from '@/components/ui/button';
import { Card, CardContent, CardHeader, CardTitle, CardDescription, CardFooter } from '@/components/ui/card';
import { Slider } from '@/components/ui/slider';
import { Input } from '@/components/ui/input';
import { Progress } from '@/components/ui/progress';
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "@/components/ui/select";
import {
  Dialog,
  DialogContent,
  DialogDescription,
  DialogHeader,
  DialogTitle,
  DialogFooter,
} from "@/components/ui/dialog";
import { Badge } from "@/components/ui/badge";
import { Label } from "@/components/ui/label";
import { Toaster, toast } from 'sonner';

import { CardGrid } from './components/CardGrid';
import { ExcelImporter } from './components/ExcelImporter';
import { exportDeckToWord } from './lib/docxGenerator';
import { generateTabooCards, getSocraticHint } from './lib/gemini';
import { CardStatus } from './types';
import { 
  auth, 
  googleProvider,
  saveSeccion,
  deleteSeccionAndStudents,
  getSecciones,
  getEstudiantes,
  saveEstudiante,
  deleteEstudiante,
  FirebaseSeccion,
  FirebaseEstudiante
} from './lib/firebase';
import { signInWithPopup, onAuthStateChanged, signOut } from 'firebase/auth';

// --- Types ---
interface CardData {
  id?: string; // ID opcional para compatibilidad con CRUD
  nivelLogro: string; // "Inicio" | "Proceso" | "Logrado" | "Destacado"
  categoria: string;
  secreto: string;
  tabu: string[];
  pista?: string;
}

interface Member {
  id: string;
  name: string;
  locked: boolean;
  role?: 'comunicador' | 'adivinador';
}

interface TeamColor {
  name: string;
  bg: string;
  text: string;
  border: string;
  lightBg: string;
  textColor: string;
}

// --- Constants ---
const TEAM_COLORS: TeamColor[] = [
  { name: "Rojo Fuego", bg: "bg-gradient-to-br from-rose-500 via-pink-500 to-rose-600 shadow-[0_6px_20px_rgba(244,63,94,0.3)] shadow-rose-500/20", text: "text-white", border: "border-rose-500/55", lightBg: "bg-rose-50/80 backdrop-blur-md", textColor: "text-rose-600 font-extrabold" },
  { name: "Azul Rayo", bg: "bg-gradient-to-br from-blue-500 via-indigo-500 to-blue-600 shadow-[0_6px_20px_rgba(59,130,246,0.3)] shadow-blue-500/20", text: "text-white", border: "border-blue-500/55", lightBg: "bg-blue-50/80 backdrop-blur-md", textColor: "text-blue-600 font-extrabold" },
  { name: "Verde Neo", bg: "bg-gradient-to-br from-emerald-400 via-teal-500 to-emerald-600 shadow-[0_6px_20px_rgba(16,185,129,0.3)] shadow-emerald-500/20", text: "text-white", border: "border-emerald-500/55", lightBg: "bg-emerald-50/80 backdrop-blur-md", textColor: "text-emerald-600 font-extrabold" },
  { name: "Naranja Fénix", bg: "bg-gradient-to-br from-amber-400 via-orange-500 to-orange-600 shadow-[0_6px_20px_rgba(249,115,22,0.3)] shadow-orange-500/20", text: "text-white", border: "border-orange-500/55", lightBg: "bg-orange-50/80 backdrop-blur-md", textColor: "text-orange-600 font-extrabold" },
  { name: "Morado Cosmos", bg: "bg-gradient-to-br from-violet-500 via-fuchsia-500 to-indigo-600 shadow-[0_6px_20px_rgba(139,92,246,0.3)] shadow-violet-500/20", text: "text-white", border: "border-violet-500/55", lightBg: "bg-violet-50/80 backdrop-blur-md", textColor: "text-violet-600 font-extrabold" },
  { name: "Celeste Glaciar", bg: "bg-gradient-to-br from-cyan-400 via-sky-500 to-blue-500 shadow-[0_6px_20px_rgba(6,182,212,0.3)] shadow-cyan-500/20", text: "text-white", border: "border-cyan-500/55", lightBg: "bg-cyan-50/80 backdrop-blur-md", textColor: "text-cyan-600 font-extrabold" },
];

const START_SLIDES = [
  {
    src: "/src/assets/images/peruvian_tabu_tv_razuri_1779245716532.png",
    bubble: "📺 PROYECCIÓN EN TV",
    title: "De espaldas a la pantalla",
    desc: "Un alumno con uniforme (camisa o blusa con corbata y falda tableada o pantalón escolar) de espaldas al televisor. La tarjeta con la palabra Tabú se proyecta desde la laptop: ¡solo tu equipo la lee!",
    hashtag: "#SanPedroDeLloc #ProyecciónTV"
  },
  {
    src: "/src/assets/images/peruvian_students_fun_razuri_1779245737446.png",
    bubble: "🗣️ PISTAS CON SEÑAS Y RISAS",
    title: "¡Explica con ingenio!",
    desc: "Compañeros luciendo blusas impecables con insignias, risas y adrenalina, gesticulan para dar pistas creativas sin decir palabras prohibidas escritas en español.",
    hashtag: "#PistasCreativas #TeamLearning"
  },
  {
    src: "/src/assets/images/peruvian_winners_podio_razuri_1779245756409.png",
    bubble: "🥇 ¡CELEBRA LA VICTORIA!",
    title: "Podio de Campeones",
    desc: "Sube a la cima del podio de ganadores en tu uniforme escolar (falda a cuadros para chicas, pantalón guinda elegante para chicos). ¡Recibe medallas y tu trofeo de campeón!",
    hashtag: "#CampeonesTabu #OrgulloEscolar"
  },
  {
    src: "/src/assets/images/peruvian_tabu_progress_razuri_1779245776635.png",
    bubble: "🎖️ PROGRESO POR NIVELES",
    title: "Reportes y Calificaciones",
    desc: "Visualiza el progreso de tus equipos en la gran pantalla. Logra los niveles de avance descritos en español, desde 'Novato' hasta coronarte como 'Leyenda Tabú'.",
    hashtag: "#NivelesYMedallas #TeamAnalytics"
  }
];

const DEFAULT_CARDS: CardData[] = [
  { nivelLogro: "Inicio", secreto: "Perro", tabu: ["Ladrar", "Mascota", "Cola", "Hueso"], categoria: "Animales", pista: "El mejor amigo del hombre que cuida la casa." },
  { nivelLogro: "Inicio", secreto: "Helado", tabu: ["Frío", "Crema", "Chocolate", "Vainilla"], categoria: "Comida", pista: "Postre refrescante de sabores cremosos." },
  { nivelLogro: "Proceso", secreto: "Teléfono", tabu: ["Llamar", "WhatsApp", "Pantalla", "Red"], categoria: "Tecnología", pista: "Dispositivo para hablar a distancia." },
  { nivelLogro: "Proceso", secreto: "Gravedad", tabu: ["Caer", "Newton", "Tierra", "Fuerza"], categoria: "Ciencia", pista: "Lo que atrae los objetos hacia el centro de la Tierra." },
  { nivelLogro: "Logrado", secreto: "Computadora", tabu: ["Mouse", "Teclado", "Internet", "RAM"], categoria: "Tecnología", pista: "Herramienta digital para procesos complejos." },
  { nivelLogro: "Logrado", secreto: "Democracia", tabu: ["Voto", "Elecciones", "Pueblo", "Libertad"], categoria: "Sociales", pista: "Sistema de gobierno basado en la voluntad popular." },
  { nivelLogro: "Destacado", secreto: "Fotosíntesis", tabu: ["Plantas", "Sol", "Luz", "Clorofila"], categoria: "Ciencia", pista: "Proceso vital de los vegetales usando la luz solar." },
  { nivelLogro: "Destacado", secreto: "Renacimiento", tabu: ["Arte", "Da Vinci", "Italia", "Historia"], categoria: "Historia", pista: "Época de florecimiento artístico tras la Edad Media." },
];

const SECTION_COLORS = [
  'indigo',
  'emerald',
  'rose',
  'amber',
  'cyan',
  'violet',
  'orange',
  'sky'
];

const COLOR_MAPPING: Record<string, { bg: string; title: string }> = {
  indigo: { bg: 'bg-indigo-500', title: 'Índigo' },
  emerald: { bg: 'bg-emerald-500', title: 'Vesmeralda' }, // 'Esmeralda' is fine but 'Verde' or similar works
  rose: { bg: 'bg-rose-500', title: 'Rosa/Rojo' },
  amber: { bg: 'bg-amber-400', title: 'Ámbar' },
  cyan: { bg: 'bg-cyan-500', title: 'Cian' },
  violet: { bg: 'bg-violet-500', title: 'Violeta' },
  orange: { bg: 'bg-orange-500', title: 'Naranja' },
  sky: { bg: 'bg-sky-400', title: 'Celeste' }
};

const getColorClasses = (colorName?: string) => {
  const norm = colorName?.toLowerCase() || 'indigo';
  switch (norm) {
    case 'emerald':
      return {
        border: 'border-emerald-250 border-l-[6px] border-l-emerald-500',
        badge: 'bg-emerald-500 text-white',
        text: 'text-emerald-900',
        bgLight: 'bg-emerald-100/85 hover:bg-emerald-150/95 hover:border-emerald-300',
        badgeDot: 'bg-emerald-600',
        textMuted: 'text-emerald-700/80',
      };
    case 'rose':
      return {
        border: 'border-rose-250 border-l-[6px] border-l-rose-500',
        badge: 'bg-rose-500 text-white',
        text: 'text-rose-900',
        bgLight: 'bg-rose-100/85 hover:bg-rose-150/95 hover:border-rose-300',
        badgeDot: 'bg-rose-600',
        textMuted: 'text-rose-700/80',
      };
    case 'amber':
      return {
        border: 'border-amber-250 border-l-[6px] border-l-amber-500',
        badge: 'bg-amber-400 text-white',
        text: 'text-amber-900',
        bgLight: 'bg-amber-100/85 hover:bg-amber-150/95 hover:border-amber-300',
        badgeDot: 'bg-amber-600',
        textMuted: 'text-amber-700/80',
      };
    case 'cyan':
      return {
        border: 'border-cyan-250 border-l-[6px] border-l-cyan-500',
        badge: 'bg-cyan-500 text-white',
        text: 'text-cyan-900',
        bgLight: 'bg-cyan-100/85 hover:bg-cyan-150/95 hover:border-cyan-300',
        badgeDot: 'bg-cyan-600',
        textMuted: 'text-cyan-700/80',
      };
    case 'violet':
      return {
        border: 'border-violet-250 border-l-[6px] border-l-violet-500',
        badge: 'bg-violet-500 text-white',
        text: 'text-violet-900',
        bgLight: 'bg-violet-100/85 hover:bg-violet-150/95 hover:border-violet-300',
        badgeDot: 'bg-violet-600',
        textMuted: 'text-violet-700/80',
      };
    case 'orange':
      return {
        border: 'border-orange-250 border-l-[6px] border-l-orange-500',
        badge: 'bg-orange-500 text-white',
        text: 'text-orange-900',
        bgLight: 'bg-orange-100/85 hover:bg-orange-150/95 hover:border-orange-300',
        badgeDot: 'bg-orange-600',
        textMuted: 'text-orange-700/80',
      };
    case 'sky':
      return {
        border: 'border-sky-250 border-l-[6px] border-l-sky-500',
        badge: 'bg-sky-400 text-white',
        text: 'text-sky-900',
        bgLight: 'bg-sky-100/85 hover:bg-sky-150/95 hover:border-sky-300',
        badgeDot: 'bg-sky-650',
        textMuted: 'text-sky-700/80',
      };
    case 'indigo':
    default:
      return {
        border: 'border-indigo-250 border-l-[6px] border-l-indigo-500',
        badge: 'bg-indigo-500 text-white',
        text: 'text-indigo-900',
        bgLight: 'bg-indigo-100/85 hover:bg-indigo-150/95 hover:border-indigo-300',
        badgeDot: 'bg-indigo-600',
        textMuted: 'text-indigo-700/80',
      };
  }
};

export default function App() {
  // --- States ---
  const [screen, setScreen] = useState<"start" | "config" | "teams" | "game" | "stats">("start");
  const [startImageIdx, setStartImageIdx] = useState(0);

  // --- Firebase Auth & Section States ---
  const [currentUser, setCurrentUser] = useState<any>(null);
  const [loadingAuth, setLoadingAuth] = useState(true);

  const [secciones, setSecciones] = useState<FirebaseSeccion[]>([]);
  const [loadingSecciones, setLoadingSecciones] = useState(false);
  
  // Managing section dialog controls:
  const [showSectionModal, setShowSectionModal] = useState(false);
  const [editingSection, setEditingSection] = useState<FirebaseSeccion | null>(null);
  const [sectionGrado, setSectionGrado] = useState("");
  const [sectionNivel, setSectionNivel] = useState("");
  const [sectionNombre, setSectionNombre] = useState("");
  const [sectionColor, setSectionColor] = useState("indigo");
  const [sectionRawStudentsText, setSectionRawStudentsText] = useState(""); // Multi-line paste roster

  // Manage student dialog controls for active editing:
  const [showEstudiantesModal, setShowEstudiantesModal] = useState(false);
  const [activeManageSection, setActiveManageSection] = useState<FirebaseSeccion | null>(null);
  const [manageStudentsList, setManageStudentsList] = useState<FirebaseEstudiante[]>([]);
  const [newStudentName, setNewStudentName] = useState("");

  // Section deletion confirmation:
  const [sectionIdToDelete, setSectionIdToDelete] = useState<string | null>(null);

  // Firebase Auth State Listener
  useEffect(() => {
    const unsubscribe = onAuthStateChanged(auth, (user) => {
      setCurrentUser(user);
      setLoadingAuth(false);
    });
    return () => unsubscribe();
  }, []);

  // Sync sections from Firestore
  const loadSeccionesFromDb = async (userId: string) => {
    setLoadingSecciones(true);
    try {
      const list = await getSecciones(userId);
      setSecciones(list);
    } catch (e) {
      console.error("Error al cargar secciones:", e);
      toast.error("Error al sincronizar tus secciones de clase.");
    } finally {
      setLoadingSecciones(false);
    }
  };

  useEffect(() => {
    if (currentUser) {
      loadSeccionesFromDb(currentUser.uid);
    } else {
      setSecciones([]);
    }
  }, [currentUser]);

  // Auth Operations
  const handleGoogleSignIn = async () => {
    try {
      const result = await signInWithPopup(auth, googleProvider);
      toast.success(`¡Bienvenido/a, ${result.user.displayName || "Docente"}!`);
    } catch (error: any) {
      console.error("Google Signin Error:", error);
      toast.error("Error o cancelación del inicio de sesión con Google.");
    }
  };

  const handleSignOut = async () => {
    try {
      await signOut(auth);
      toast.info("Sesión cerrada.");
    } catch (error: any) {
      console.error("Signout Error:", error);
      toast.error("Error al cerrar sesión.");
    }
  };

  // Section Operations
  const handleCreateOrEditSection = async () => {
    if (!currentUser) return;
    if (!sectionGrado.trim() || !sectionNivel.trim() || !sectionNombre.trim()) {
      toast.error("Por favor completa Grado, Nivel e Identificador o Letra.");
      return;
    }

    const secId = editingSection?.id || Math.random().toString(36).substring(2, 11);
    const newSec: FirebaseSeccion = {
      id: secId,
      userId: currentUser.uid,
      grado: sectionGrado.trim(),
      nivel: sectionNivel.trim(),
      nombre: sectionNombre.trim(),
      color: sectionColor
    };

    const isEdit = !!editingSection;
    const studentsRaw = sectionRawStudentsText;

    // Immediately close modal and reset inputs to provide instant UI feedback
    setShowSectionModal(false);
    setSectionGrado("");
    setSectionNivel("");
    setSectionNombre("");
    setSectionColor("indigo");
    setSectionRawStudentsText("");
    setEditingSection(null);

    const toastId = toast.loading(isEdit ? "Guardando cambios de sección..." : "Guardando nueva sección y alumnos en la nube...");

    try {
      // Create or update section
      await saveSeccion(newSec, isEdit);

      // Save initial batch of student names if registering a new section
      if (!isEdit && studentsRaw.trim()) {
        const names = studentsRaw.split('\n')
          .map(n => n.trim())
          .filter(n => n !== "");
        
        for (const name of names) {
          const estud: FirebaseEstudiante = {
            id: Math.random().toString(36).substring(2, 9),
            name,
            role: 'comunicador',
            locked: false
          };
          await saveEstudiante(secId, estud);
        }
      }

      toast.success(isEdit ? "Sección actualizada correctamente." : "Sección y alumnos creados correctamente.", { id: toastId });
      
      // Refresh list
      await loadSeccionesFromDb(currentUser.uid);
    } catch (e) {
      console.error(e);
      toast.error("Error al guardar la sección de clase.", { id: toastId });
    }
  };

  // Student Operations
  const openManageStudents = async (parentSec: FirebaseSeccion) => {
    setActiveManageSection(parentSec);
    setNewStudentName("");
    try {
      const list = await getEstudiantes(parentSec.id);
      setManageStudentsList(list.sort((a, b) => a.name.localeCompare(b.name)));
      setShowEstudiantesModal(true);
    } catch (e) {
      console.error(e);
      toast.error("Error al cargar estudiantes.");
    }
  };

  const handleAddStudentToSection = async () => {
    if (!activeManageSection || !newStudentName.trim()) {
      toast.error("Ingrese el nombre del estudiante.");
      return;
    }

    const newEstud: FirebaseEstudiante = {
      id: Math.random().toString(36).substring(2, 9),
      name: newStudentName.trim(),
      role: 'comunicador',
      locked: false
    };

    try {
      await saveEstudiante(activeManageSection.id, newEstud);
      setNewStudentName("");
      // Reload list
      const updatedList = await getEstudiantes(activeManageSection.id);
      setManageStudentsList(updatedList.sort((a, b) => a.name.localeCompare(b.name)));
      toast.success("Estudiante agregado correctamente.");
      
      // Update our main list counts
      if (currentUser) {
        await loadSeccionesFromDb(currentUser.uid);
      }
    } catch (e) {
      console.error(e);
      toast.error("Error al agregar el alumno.");
    }
  };

  const handleDeleteStudentFromSection = async (estudId: string) => {
    if (!activeManageSection) return;
    try {
      await deleteEstudiante(activeManageSection.id, estudId);
      // Reload list
      const updatedList = await getEstudiantes(activeManageSection.id);
      setManageStudentsList(updatedList.sort((a, b) => a.name.localeCompare(b.name)));
      toast.success("Estudiante retirado/eliminado.");
    } catch (e) {
      console.error(e);
      toast.error("Error al eliminar el estudiante.");
    }
  };

  const handleDeleteSection = (secId: string) => {
    setSectionIdToDelete(secId);
  };

  // Set selected section & populate teams active players
  const handleStartGameWithSection = async (sec: FirebaseSeccion) => {
    try {
      const students = await getEstudiantes(sec.id);
      if (students.length === 0) {
        toast.error("Esta sección aún no tiene estudiantes registrados. Agrega estudiantes primero.");
        return;
      }

      // 1. Populate standard studentList text string
      const namesString = students.map(s => s.name).join('\n');
      setStudentList(namesString);

      // 2. Prep structured game ready members
      const gameMembers: Member[] = students.map(s => ({
        id: s.id,
        name: s.name,
        locked: false,
        role: 'comunicador'
      }));

      // 3. Clear/Restart and spread evenly to active teams
      const nextMembers: Member[][] = Array.from({ length: numTeams }, () => []);
      let currentT = 0;
      gameMembers.forEach(member => {
        nextMembers[currentT].push(member);
        currentT = (currentT + 1) % numTeams;
      });

      setTeamMembers(nextMembers);
      toast.success(`Estudiantes de la sección "${sec.grado} - ${sec.nombre}" de nivel ${sec.nivel} cargados.`);
      
      // Switch screen to team configurations
      setScreen("teams");
    } catch (e) {
      console.error(e);
      toast.error("Error al cargar la sección para jugar hoy.");
    }
  };

  
  // Config States
  const [numTeams, setNumTeams] = useState(2);
  const [studentList, setStudentList] = useState("");
  const [teamMembers, setTeamMembers] = useState<Member[][]>([[], [], [], [], [], []]);
  const [draggingMember, setDraggingMember] = useState<{tIdx: number, mIdx: number} | null>(null);
  const [totalRounds, setTotalRounds] = useState(3);
  const [passesPerTurn, setPassesPerTurn] = useState(2);
  const [timePerTurn, setTimePerTurn] = useState(60);

  // Helper for safe state updates
  const safeUpdate = (setter: React.Dispatch<React.SetStateAction<number>>, val: number, min: number, max: number) => {
    const num = Number(val);
    if (isNaN(num)) return;
    setter(Math.min(max, Math.max(min, num)));
  };
  const [selectedCategory, setSelectedCategory] = useState<string>("all");
  const [selectedNivelLogro, setSelectedNivelLogro] = useState<string>("all");
  const [aiTopic, setAiTopic] = useState("");
  const [isGenerating, setIsGenerating] = useState(false);
  const [configTab, setConfigTab] = useState<"general" | "deck">("general");
  const [isExcelImportActive, setIsExcelImportActive] = useState(false);
  const [exportColumns, setExportColumns] = useState<number>(3);
  const [downloadFileName, setDownloadFileName] = useState<string>("tarjetas_tabu_editables");

  // Game Logic States
  const [teamScores, setTeamScores] = useState<number[]>([]);
  const [turnCounter, setTurnCounter] = useState(0);
  const [currentCard, setCurrentCard] = useState<CardData | null>(null);
  const [teamDiffCycles, setTeamDiffCycles] = useState<number[]>([]);
  const [teamCommunicatorCount, setTeamCommunicatorCount] = useState<number[]>([1, 1, 1, 1, 1, 1]);
  const [activeCommunicators, setActiveCommunicators] = useState<Member[]>([]);

  // Derived Values
  const currentTeamIndex = turnCounter % (teamScores.length || 1);
  
  // A team's "current round" is how many times they've played, 
  // plus the round they are currently in or about to play.
  const getTeamRound = (idx: number) => {
    const n = teamScores.length || 1;
    return Math.floor(turnCounter / n) + (idx < (turnCounter % n) ? 1 : 0);
  };
  const [showTurnModal, setShowTurnModal] = useState(false);
  const [usedWords, setUsedWords] = useState<string[]>([]);
  const [passesUsedInTurn, setPassesUsedInTurn] = useState(0);
  const [timeLeft, setTimeLeft] = useState(60);
  const [isTimerActive, setIsTimerActive] = useState(false);
  const [loadedCards, setLoadedCards] = useState<CardData[]>([]);
  const [socraticHint, setSocraticHint] = useState<string | null>(null);
  const [isHintsVisible, setIsHintsVisible] = useState(true);
  const [isHintLoading, setIsHintLoading] = useState(false);

  const timerRef = useRef<NodeJS.Timeout | null>(null);

  // --- Helpers ---
  const getAvailableCards = () => {
    const all = loadedCards.length > 0 ? loadedCards : DEFAULT_CARDS;
    let filtered = all;
    if (selectedCategory !== "all") {
      filtered = filtered.filter(c => c.categoria === selectedCategory);
    }
    if (selectedNivelLogro !== "all") {
      filtered = filtered.filter(c => c.nivelLogro === selectedNivelLogro);
    }
    return filtered;
  };

  const getCategories = () => {
    const all = loadedCards.length > 0 ? loadedCards : DEFAULT_CARDS;
    return Array.from(new Set(all.map(c => c.categoria)));
  };

  // --- Handlers ---
  const updateMember = (teamIdx: number, memberIdx: number, name: string) => {
    setTeamMembers(prev => {
      const next = [...prev];
      next[teamIdx] = [...next[teamIdx]];
      next[teamIdx][memberIdx] = { ...next[teamIdx][memberIdx], name };
      return next;
    });
  };

  const toggleMemberLock = (teamIdx: number, memberIdx: number) => {
    setTeamMembers(prev => {
      const next = [...prev];
      if (!next[teamIdx]) next[teamIdx] = [];
      next[teamIdx] = [...next[teamIdx]];
      if (next[teamIdx][memberIdx]) {
        next[teamIdx][memberIdx] = { ...next[teamIdx][memberIdx], locked: !next[teamIdx][memberIdx].locked };
      }
      return next;
    });
  };

  const toggleMemberRole = (teamIdx: number, memberIdx: number) => {
    setTeamMembers(prev => {
      const next = [...prev];
      if (!next[teamIdx]) next[teamIdx] = [];
      next[teamIdx] = [...next[teamIdx]];
      if (next[teamIdx][memberIdx]) {
        const currentRole = next[teamIdx][memberIdx].role || 'comunicador';
        const nextRole = currentRole === 'comunicador' ? 'adivinador' : 'comunicador';
        next[teamIdx][memberIdx] = { ...next[teamIdx][memberIdx], role: nextRole };
      }
      return next;
    });
  };

  const addMember = (teamIdx: number) => {
    if ((teamMembers[teamIdx]?.length || 0) >= 8) {
      toast.error("Máximo 8 alumnos por equipo");
      return;
    }
    setTeamMembers(prev => {
      const next = [...prev];
      if (!next[teamIdx]) next[teamIdx] = [];
      next[teamIdx] = [...next[teamIdx], { 
        id: Math.random().toString(36).substr(2, 9), 
        name: `Alumno ${(next[teamIdx]?.length || 0) + 1}`, 
        locked: false,
        role: 'comunicador'
      }];
      return next;
    });
  };

  const removeMember = (teamIdx: number, memberIdx: number) => {
    setTeamMembers(prev => {
      const next = [...prev];
      next[teamIdx] = next[teamIdx].filter((_, i) => i !== memberIdx);
      return next;
    });
  };

  const shuffleTeams = () => {
    // Collect all members currently in the systems
    const allExistingMembers = teamMembers.flat();
    
    // If we have some members already, we might want to use them.
    // If not, we use the studentList.
    let pool: Member[] = [];
    
    const listNames = studentList.split('\n').map(n => n.trim()).filter(n => n !== "");
    
    if (allExistingMembers.length > 0) {
      // Use existing members but respect locks
      pool = [...allExistingMembers];
    } else if (listNames.length > 0) {
      // Use list names
      pool = listNames.map(name => ({
        id: Math.random().toString(36).substr(2, 9),
        name,
        locked: false,
        role: 'comunicador'
      }));
    } else {
      toast.error("Ingresa una lista de alumnos o agrega integrantes primero");
      return;
    }

    // Separate locked members
    const nextMembers: Member[][] = Array.from({ length: numTeams }, () => []);
    
    // Keep locked members in their current teams (if they were already there)
    // Actually, it's better to process team-by-team to keep their positions relative to the team
    const unlockedToShuffle: Member[] = [];
    
    teamMembers.forEach((team, tIdx) => {
      if (tIdx >= numTeams) {
          // If team index is now out of range, all its members become unlocked to shuffle
          unlockedToShuffle.push(...team.map(m => ({ ...m, locked: false })));
          return;
      }
      team.forEach(member => {
        if (member.locked) {
          nextMembers[tIdx].push(member);
        } else {
          unlockedToShuffle.push(member);
        }
      });
    });

    // If we are starting from studentList, all are unlocked
    if (allExistingMembers.length === 0 && listNames.length > 0) {
        unlockedToShuffle.length = 0;
        unlockedToShuffle.push(...pool);
    }

    // Shuffle the unlocked pool
    const shuffled = [...unlockedToShuffle].sort(() => Math.random() - 0.5);
    
    // Redistribute shuffled members to fill teams up to a balanced capacity
    // To be fair, we distribute round-robin to teams that have space
    let currentTeam = 0;
    shuffled.forEach(member => {
      // Find the next team that has space (max 8)
      let found = false;
      let attempts = 0;
      while (!found && attempts < numTeams) {
        const targetIdx = (currentTeam + attempts) % numTeams;
        if (nextMembers[targetIdx].length < 8) {
          nextMembers[targetIdx].push(member);
          currentTeam = (targetIdx + 1) % numTeams;
          found = true;
        }
        attempts++;
      }
    });
    
    setTeamMembers(nextMembers);
    toast.success("¡Equipos reorganizados respetando bloqueos!");
  };

  const clearTeamsAndStudents = () => {
    setStudentList("");
    setTeamMembers(Array.from({ length: numTeams }, () => []));
    toast.success("Organización de integrantes limpia.");
  };

  const handleDragStart = (tIdx: number, mIdx: number) => {
    setDraggingMember({ tIdx, mIdx });
  };

  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
  };

  const handleDrop = (targetTIdx: number) => {
    if (draggingMember === null) return;
    const { tIdx: sourceTIdx, mIdx: sourceMIdx } = draggingMember;
    
    if (sourceTIdx === targetTIdx) {
      setDraggingMember(null);
      return;
    }

    setTeamMembers(prev => {
      const next = [...prev.map(t => [...t])];
      const member = next[sourceTIdx][sourceMIdx];
      
      if (next[targetTIdx].length >= 8) {
        toast.error("El equipo destino está lleno (Máx 8)");
        return prev;
      }

      // Remove from source
      next[sourceTIdx].splice(sourceMIdx, 1);
      // Add to target
      next[targetTIdx].push(member);
      
      return next;
    });
    setDraggingMember(null);
  };

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const data = new Uint8Array(evt.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 }) as any[][];

        // Format: Nivel de Logro, Category, Secret, Taboo1, Taboo2, Taboo3, Taboo4, Hint
        const newCards: CardData[] = rows.slice(1)
          .filter(row => row[2] && row[1]) // Ensure secret (row[2]) and category (row[1]) are defined
          .map(row => {
            let nl = String(row[0] || "").trim();
            if (!nl) {
              nl = "Inicio"; // Default fallback
            }
            return {
              nivelLogro: nl,
              categoria: String(row[1]).trim(),
              secreto: String(row[2]).trim(),
              tabu: [
                String(row[3] || "").trim(),
                String(row[4] || "").trim(),
                String(row[5] || "").trim(),
                String(row[6] || "").trim(),
              ].filter(Boolean),
              pista: row[7] ? String(row[7]).trim() : undefined
            };
          });

        if (newCards.length === 0) throw new Error("No se encontraron tarjetas válidas.");
        
        setLoadedCards(newCards);
        toast.success(`${newCards.length} tarjetas cargadas con éxito.`);
      } catch (err) {
        toast.error("Error al procesar el archivo Excel.");
      }
    };
    reader.readAsArrayBuffer(file);
    e.target.value = ''; // Reset input
  };

  const handleStudentsUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const data = new Uint8Array(evt.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 }) as any[][];

        // Format: Each column is a team. Rows are names. Max columns = 6 (numTeams)
        const nextMembers: Member[][] = [[], [], [], [], [], []];
        let maxColumnWithContent = 0;
        
        // Find max content across rows
        rows.forEach((row, ri) => {
          if (ri === 0) return; // Skip header
          row.forEach((cell, ci) => {
            if (ci < 6 && cell) {
              const name = String(cell).trim();
              if (name && nextMembers[ci].length < 8) {
                nextMembers[ci].push({
                  id: Math.random().toString(36).substr(2, 9),
                  name,
                  locked: false,
                  role: 'comunicador'
                });
                maxColumnWithContent = Math.max(maxColumnWithContent, ci);
              }
            }
          });
        });

        // Set the active number of teams matching the imported file
        const detectedTeams = Math.max(2, Math.min(6, maxColumnWithContent + 1));
        setNumTeams(detectedTeams);

        setTeamMembers(nextMembers);
        toast.success(`Listado de alumnos cargado con ${detectedTeams} equipos.`);
      } catch (err) {
        toast.error("Error al procesar el listado de alumnos.");
      }
    };
    reader.readAsArrayBuffer(file);
    e.target.value = ''; // Reset input
  };

  const downloadTemplate = (type: 'cards' | 'students') => {
    let data: any[][] = [];
    let filename = "";

    if (type === 'cards') {
      data = [
        ["Nivel de Logro", "Categoría", "Palabra Secreta", "Tabú 1", "Tabú 2", "Tabú 3", "Tabú 4", "Pista (Opcional)"],
        ["Inicio", "Energía y Combustibles", "Energía", "Capacidad", "Cuerpos", "Cambios", "Trabajo", "Propiedad de la materia para realizar transformaciones."],
        ["Proceso", "Energía y Combustibles", "Turbina", "Hélice", "Generador", "Rotación", "Vapor", "Rueda giratoria movida por un fluido para producir trabajo mecánico."],
        ["Logrado", "Energía y Combustibles", "Átomo", "Proceso", "Análisis", "Variable", "Sistema", "Ficha complementaria del bloque temático: Energía y Combustibles."],
        ["Destacado", "Energía y Combustibles", "Enlace Químico", "Proceso", "Análisis", "Variable", "Sistema", "Ficha complementaria del bloque temático: Energía y Combustibles."]
      ];
      filename = "TabuEdu_Plantilla_Tarjetas.xlsx";
    } else {
      data = [
        ["Equipo 1", "Equipo 2", "Equipo 3", "Equipo 4", "Equipo 5", "Equipo 6"],
        ["Juan Pérez", "Ana García", "Carlos Ruíz", "", "", ""],
        ["María Luz", "José Torres", "Rosa Vidal", "", "", ""]
      ];
      filename = "TabuEdu_Plantilla_Alumnos.xlsx";
    }

    const ws = XLSX.utils.aoa_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Hoja1");
    XLSX.writeFile(wb, filename);
  };

  const downloadCurrentTeams = () => {
    const headers = Array.from({ length: numTeams }, (_, i) => TEAM_COLORS[i].name);
    const maxMembers = Math.max(...teamMembers.slice(0, numTeams).map(t => t?.length || 0), 0);
    
    const rows = [headers];
    for (let r = 0; r < maxMembers; r++) {
      const row = Array.from({ length: numTeams }, (_, c) => teamMembers[c]?.[r]?.name || "");
      rows.push(row);
    }
    
    const filename = `TabuEdu_Distribucion_Equipos.xlsx`;
    const ws = XLSX.utils.aoa_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Equipos");
    XLSX.writeFile(wb, filename);
    toast.success("Distribución actual por equipos descargada.");
  };

  const generateWithAI = async () => {
    if (!aiTopic.trim()) {
      toast.error("Por favor ingresa un tema.");
      return;
    }

    setIsGenerating(true);
    const toastId = toast.info(`Generando más de 200 tarjetas para "${aiTopic}" en lotes con IA... Por favor espera.`, { duration: 15000 });
    
    try {
      const batches = [
        `${aiTopic} (Parte 1: Términos fundamentales, conceptos iniciales e historia)`,
        `${aiTopic} (Parte 2: Teorías, métodos, aplicaciones y procesos de aprendizaje)`,
        `${aiTopic} (Parte 3: Conceptos técnicos de nivel intermedio y avanzado)`,
        `${aiTopic} (Parte 4: Detalles especializados, herramientas complejas y casos prácticos)`
      ];

      const results: CardData[][] = [];
      const sleep = (ms: number) => new Promise((resolve) => setTimeout(resolve, ms));

      for (let i = 0; i < batches.length; i++) {
        const batchName = batches[i];
        toast.info(`[Lote ${i + 1}/4] Generando tarjetas sobre "${aiTopic}"... Por favor, espera.`, { 
          id: toastId, 
          duration: 30000 
        });
        
        let attempts = 0;
        const maxAttempts = 3;
        let success = false;
        let batchCards: CardData[] = [];
        
        while (attempts < maxAttempts && !success) {
          try {
            batchCards = await generateTabooCards(batchName, 55);
            success = true;
          } catch (err: any) {
            attempts++;
            const errMsg = String(err.message || "").toLowerCase();
            const isRateLimit = errMsg.includes("429") || 
                                errMsg.includes("quota") || 
                                errMsg.includes("exhausted") || 
                                errMsg.includes("limite") || 
                                errMsg.includes("retry") ||
                                errMsg.includes("rate");
            
            if (attempts < maxAttempts && isRateLimit) {
              // Increase wait time per retry
              const waitTime = 15000 + attempts * 8000;
              toast.info(`[Lote ${i + 1}/4] Límite de cuota detectado. Reintentando en ${Math.round(waitTime / 1000)}s... (Intento ${attempts}/${maxAttempts})`, { 
                id: toastId, 
                duration: waitTime + 2000 
              });
              await sleep(waitTime);
            } else {
              throw err;
            }
          }
        }
        results.push(batchCards);
        
        // Brief pause between successful batches to respect Rate Limits smoothly
        if (i < batches.length - 1) {
          await sleep(2500);
        }
      }

      const allGenerated = results.flat();

      if (allGenerated.length === 0) {
        throw new Error("No se pudo generar ninguna tarjeta.");
      }

      setLoadedCards(prev => {
        // Collect default secrets to identify preset cards
        const defaultSecrets = new Set(DEFAULT_CARDS.map(c => c.secreto.toLowerCase().trim()));
        
        // Filter out preset/default cards from existing load
        const existingList = (prev.length > 0 ? prev : []).filter(c => {
          const isDefaultId = c.id && c.id.toString().startsWith("default-");
          const isDefaultSecret = defaultSecrets.has(c.secreto.toLowerCase().trim());
          return !isDefaultId && !isDefaultSecret;
        });
        
        const removedDefaultCount = (prev.length > 0 ? prev.length : DEFAULT_CARDS.length) - existingList.length;
        const existingSecretos = new Set(existingList.map(c => c.secreto.toLowerCase().trim()));
        
        const addedCards: CardData[] = [];
        allGenerated.forEach(card => {
          const sec = card.secreto.toLowerCase().trim();
          if (!existingSecretos.has(sec)) {
            existingSecretos.add(sec);
            addedCards.push(card);
          }
        });

        const newTotalList = [...existingList, ...addedCards];

        toast.dismiss(toastId);
        
        let msg = `Se generaron y acumularon ${addedCards.length} nuevas tarjetas sobre "${aiTopic}". ¡Total en mazo: ${newTotalList.length}!`;
        if (removedDefaultCount > 0) {
          msg += ` Se eliminaron automáticamente ${removedDefaultCount} tarjetas predeterminadas.`;
        }
        toast.success(msg);
        return newTotalList;
      });

      setAiTopic("");
    } catch (err: any) {
      console.error(err);
      toast.dismiss(toastId);
      toast.error(err.message || "Error al generar tarjetas con IA.");
    } finally {
      setIsGenerating(false);
    }
  };

  const proceedToTeams = () => {
    const available = getAvailableCards();
    if (available.length < 5) {
      toast.error("Necesitas al menos 5 tarjetas para jugar.");
      return;
    }
    setScreen("teams");
  };

  const startGame = () => {
    // Ensure teamMembers has enough arrays for the number of teams
    setTeamMembers(prev => {
      const sanitized = [...prev];
      for (let i = 0; i < numTeams; i++) {
        if (!sanitized[i]) sanitized[i] = [];
      }
      return sanitized;
    });

    setTeamScores(Array(numTeams).fill(0));
    setTurnCounter(0);
    setUsedWords([]);
    setTeamDiffCycles(Array(numTeams).fill(0)); // Start each team at difficulty index 0 (Inicio)
    setTimeLeft(timePerTurn); // Ensure timer is ready
    setScreen("game");
    setShowTurnModal(true);
  };

  const startTurn = () => {
    setShowTurnModal(false);
    setPassesUsedInTurn(0);
    setTimeLeft(timePerTurn); // Reset timer at the exact moment turn starts
    setIsTimerActive(true);
    setSocraticHint(null);
    nextCard();
  };

  const nextCard = (justUsedSecret?: string, advanceCycle: boolean = false) => {
    const available = getAvailableCards().filter(c => {
      const isUsed = usedWords.includes(c.secreto) || (justUsedSecret ? c.secreto === justUsedSecret : false);
      return !isUsed;
    });

    if (available.length === 0) {
      toast.info("¡Se acabaron las palabras!");
      endGame();
      return;
    }

    // Determine target index
    let currentIdx = teamDiffCycles[currentTeamIndex] !== undefined ? teamDiffCycles[currentTeamIndex] : 0;
    if (advanceCycle) {
      currentIdx = (currentIdx + 1) % 4;
      // Update state for persistence
      setTeamDiffCycles(prev => {
        const next = [...prev];
        next[currentTeamIndex] = currentIdx;
        return next;
      });
    }

    const DIFF_SEQUENCE: CardStatus[] = ["Inicio", "Proceso", "Logrado", "Destacado"];
    let targetDifficulty = DIFF_SEQUENCE[currentIdx];
    let filteredByDiff = available.filter(c => c.nivelLogro === targetDifficulty);

    // If no unused cards of the target difficulty are left, try to find the next difficulty in the sequence that has cards
    if (filteredByDiff.length === 0 && available.length > 0) {
      let nextCycleIdx = (currentIdx + 1) % 4;
      let found = false;
      for (let i = 0; i < 4; i++) {
        const testDiff = DIFF_SEQUENCE[nextCycleIdx];
        const testFiltered = available.filter(c => c.nivelLogro === testDiff);
        if (testFiltered.length > 0) {
          targetDifficulty = testDiff;
          filteredByDiff = testFiltered;
          currentIdx = nextCycleIdx;
          setTeamDiffCycles(prev => {
            const next = [...prev];
            next[currentTeamIndex] = nextCycleIdx;
            return next;
          });
          found = true;
          break;
        }
        nextCycleIdx = (nextCycleIdx + 1) % 4;
      }
    }

    const finalPool = filteredByDiff.length > 0 ? filteredByDiff : available;
    const randomCard = finalPool[Math.floor(Math.random() * finalPool.length)];
    setCurrentCard(randomCard);
    
    // Auto-display hint if available, otherwise fetch it
    if (randomCard.pista) {
      setSocraticHint(randomCard.pista);
    } else {
      setSocraticHint(null);
      // Trigger auto-fetch for cards without hints
      fetchSocraticHint(randomCard.secreto);
    }
  };

  const fetchSocraticHint = async (word?: string) => {
    const targetWord = word || currentCard?.secreto;
    if (!targetWord || isHintLoading) return;
    setIsHintLoading(true);
    try {
      const hint = await getSocraticHint(targetWord);
      setSocraticHint(hint);
    } catch (error) {
      console.error(error);
      // Silent error if auto-fetching
    } finally {
      setIsHintLoading(false);
    }
  };

  const getGuesser = () => {
    const n = teamScores.length || 1;
    const members = teamMembers[currentTeamIndex];
    if (members && Array.isArray(members) && members.length > 0) {
      const roundIdx = Math.floor(turnCounter / n);
      return members[roundIdx % members.length]?.name || "Participante";
    }
    return "Participante";
  };

  const handleCorrect = () => {
    if (!currentCard) return;
    const secret = currentCard.secreto;
    setTeamScores(prev => {
      const next = [...prev];
      next[currentTeamIndex] = (next[currentTeamIndex] || 0) + 1;
      return next;
    });
    setUsedWords(prev => [...prev, secret]);
    nextCard(secret, true); // Advance difficulties cycle!
    toast.success("+1 Punto", { duration: 1000 });
  };

  const handleTaboo = () => {
    if (!currentCard) return;
    const secret = currentCard.secreto;
    setTeamScores(prev => {
      const next = [...prev];
      next[currentTeamIndex] = (next[currentTeamIndex] || 0) - 1;
      return next;
    });
    setUsedWords(prev => [...prev, secret]);
    nextCard(secret, true); // Advance difficulties cycle!
    toast.error("¡Tabú! -1 Punto", { duration: 1000 });
  };

  const handlePass = () => {
    if (passesUsedInTurn >= passesPerTurn) {
      toast.warning("No te quedan más pases.");
      return;
    }
    setPassesUsedInTurn(prev => prev + 1);
    nextCard(undefined, true); // Advance difficulties cycle!
  };

  const endTurn = React.useCallback(() => {
    if (!isTimerActive) return;
    
    setIsTimerActive(false);
    // No need to setTimeLeft here as we set it in startTurn
    setPassesUsedInTurn(0);
    
    setTurnCounter(prev => {
      const next = prev + 1;
      const n = teamScores.length || 1;
      const totalTurns = n * totalRounds;
      
      if (next >= totalTurns) {
        setScreen("stats");
        return next;
      } else {
        setShowTurnModal(true);
        return next;
      }
    });
  }, [isTimerActive, teamScores.length, totalRounds]);

  const endGame = () => {
    setIsTimerActive(false);
    setScreen("stats");
  };

  const getGuesserNames = () => {
    const list = teamMembers[currentTeamIndex] || [];
    if (list.length === 0) return "Cualquiera";
    
    // Explicit adivinadores
    const explicitAdvs = list.filter(m => m.role === 'adivinador');
    if (explicitAdvs.length > 0) {
      return explicitAdvs.map(m => m.name).join(', ');
    }
    
    // Rest of the team who are not communicators this turn
    const activeCommNames = activeCommunicators.map(c => c.name);
    const rest = list.filter(m => !activeCommNames.includes(m.name));
    if (rest.length > 0) {
      return rest.map(m => m.name).join(', ');
    }
    
    return "Todo el equipo";
  };

  // Select active communicators randomly for the current team once per turn
  useEffect(() => {
    if (screen !== "game") return;
    
    const teamIdx = currentTeamIndex;
    const list = teamMembers[teamIdx] || [];
    if (list.length === 0) {
      setActiveCommunicators([]);
      return;
    }

    // Candidates: pool of members designated as 'comunicador'
    let candidates = list.filter(m => (m.role || 'comunicador') === 'comunicador');
    if (candidates.length === 0) {
      // Fallback
      candidates = list;
    }

    const desiredCount = teamCommunicatorCount[teamIdx] || 1;
    const selected: Member[] = [];
    const poolCopy = [...candidates];
    
    // Pick from communicator pool
    const limit = Math.min(poolCopy.length, desiredCount);
    for (let i = 0; i < limit; i++) {
      const idx = Math.floor(Math.random() * poolCopy.length);
      selected.push(poolCopy[idx]);
      poolCopy.splice(idx, 1);
    }

    // Fill the gap from remaining team members if we need 2 but have 1
    if (selected.length < desiredCount && list.length > selected.length) {
      const restCandidates = list.filter(m => !selected.some(s => s.id === m.id));
      const needed = desiredCount - selected.length;
      const extraLimit = Math.min(restCandidates.length, needed);
      for (let i = 0; i < extraLimit; i++) {
        const idx = Math.floor(Math.random() * restCandidates.length);
        selected.push(restCandidates[idx]);
        restCandidates.splice(idx, 1);
      }
    }

    setActiveCommunicators(selected);
  }, [turnCounter, currentTeamIndex, screen, teamMembers, teamCommunicatorCount]);

  // --- Effects ---
  // Ensure teamMembers always has at least numTeams sub-arrays to avoid out-of-bounds rendering crashes
  useEffect(() => {
    setTeamMembers(prev => {
      if (prev.length >= numTeams) return prev;
      const next = [...prev];
      while (next.length < numTeams) {
        next.push([]);
      }
      return next;
    });
  }, [numTeams]);

  useEffect(() => {
    let interval: NodeJS.Timeout | null = null;
    
    if (isTimerActive && timeLeft > 0) {
      interval = setInterval(() => {
        setTimeLeft(prev => Math.max(0, prev - 1));
      }, 1000);
    }

    return () => {
      if (interval) clearInterval(interval);
    };
  }, [isTimerActive, timeLeft > 0]); // Re-run when active state or "has time" changes

  // Single source of truth for turn ending
  useEffect(() => {
    if (isTimerActive && timeLeft <= 0) {
      endTurn();
      toast.info("¡Tiempo agotado!");
    }
  }, [timeLeft, isTimerActive, endTurn]);

  // Slideshow interval for youth gallery
  useEffect(() => {
    if (screen !== "start") return;
    const interval = setInterval(() => {
      setStartImageIdx(prev => (prev + 1) % START_SLIDES.length);
    }, 6000);
    return () => clearInterval(interval);
  }, [screen]);

  // --- Render Helpers ---
  const renderStart = () => {
    // Elegant letter-by-letter animation for energetic youth engagement
    const renderAnimatedLetters = (word: string, gradientClass: string) => {
      return word.split("").map((char, index) => (
        <motion.span
          key={index}
          className={`inline-block font-black uppercase tracking-tighter ${gradientClass}`}
          animate={{
            y: [0, -12, 0],
            rotate: [0, 8, -8, 0],
            scale: [1, 1.15, 0.9, 1]
          }}
          transition={{
            duration: 1.8,
            repeat: Infinity,
            ease: "easeInOut",
            delay: index * 0.1,
          }}
          style={{ originX: 0.5, originY: 0.8 }}
        >
          {char}
        </motion.span>
      ));
    };

    return (
      <motion.div 
        initial={{ opacity: 0, y: 15 }}
        animate={{ opacity: 1, y: 0 }}
        exit={{ opacity: 0, y: -15 }}
        className="grid grid-cols-1 md:grid-cols-12 max-w-[var(--container-xl)] w-full mx-auto p-[var(--fluid-space-md)] md:p-[var(--fluid-space-lg)] h-full justify-center items-center overflow-y-auto no-scrollbar"
        style={{ 
          gap: 'var(--fluid-space-md)',
          maxWidth: 'min(90rem, 100%)'
        }}
      >
        {/* Left Side: Game Pitch, Animated Title and Action Button */}
        <div className="md:col-span-6 flex flex-col justify-center text-center md:text-left items-center md:items-start" style={{ gap: 'var(--fluid-space-md)' }}>
          
          <div className="flex flex-col items-center md:items-start" style={{ gap: 'var(--fluid-space-xs)' }}>
            {/* Holographic Glowing Launchpad Badge */}
            <div className="inline-flex items-center gap-2 bg-gradient-to-r from-blue-500/10 to-indigo-500/10 p-2 px-4 rounded-full border border-blue-500/20 shadow-md">
              <Sparkles className="text-blue-500 animate-spin animate-none" style={{ width: '16px', height: '16px' }} />
              <span className="font-extrabold text-blue-600 tracking-wide uppercase" style={{ fontSize: '10px' }}>¡MODO JUVENIL ACTIVADO! ⚡🔥</span>
            </div>
            
            {/* Elite Animated Typography Wave */}
            <h1 
              className="font-black text-slate-900 leading-none pb-1 flex items-center justify-center md:justify-start gap-1"
              style={{ fontSize: 'calc(var(--fluid-text-hero) * 1.05)' }}
            >
              <span className="flex">{renderAnimatedLetters("TABÚ", "bg-gradient-to-r from-blue-600 via-indigo-600 to-purple-600 bg-clip-text text-transparent drop-shadow-sm")}</span>
              <span className="flex ml-2">{renderAnimatedLetters("EDU", "text-slate-800 font-display drop-shadow-sm")}</span>
            </h1>

            <p 
              className="text-slate-500 font-extrabold uppercase tracking-wide text-balance md:text-left text-center leading-snug"
              style={{ fontSize: 'var(--fluid-text-sm)', maxWidth: 'min(36rem, 100%)' }}
            >
              ¡El juego de palabras prohibidas que enciende la chispa en el aula o con amigos! Desafía tu mente, ríe sin parar y expande tu vocabulario superando el cronómetro.
            </p>
          </div>

          {/* Gamified Tactile Empezar Button Row */}
          <div className="w-full flex justify-center md:justify-start" style={{ maxWidth: 'min(30rem, 100%)' }}>
            <motion.div className="w-full" whileHover={{ scale: 1.05 }} whileTap={{ scale: 0.95 }}>
              <Button 
                size="lg" 
                className="w-full font-black uppercase tracking-widest bg-gradient-to-r from-blue-600 via-indigo-500 to-violet-600 hover:from-blue-700 hover:to-violet-700 text-white border-none shadow-[0_20px_45px_-12px_rgba(79,70,229,0.35)] transition-all rounded-[1.5rem]" 
                style={{ height: 'var(--fluid-space-xl)', fontSize: 'var(--fluid-text-base)' }}
                onClick={() => setScreen("config")}
              >
                <Play 
                  className="mr-3 fill-current animate-pulse text-white" 
                  style={{ width: 'var(--fluid-space-sm)', height: 'var(--fluid-space-sm)' }}
                /> COMENZAR PARTIDA
              </Button>
            </motion.div>
          </div>

          {/* CONTROL DE SECCIONES CON GOOGLE AUTH (VIRTUAL ROSTER MODULE) */}
          <div className="w-full border-t border-slate-200/60 pt-5 mt-1 flex flex-col gap-4 text-left" style={{ maxWidth: 'min(54rem, 100%)' }}>
            {!currentUser ? (
              <div className="bg-slate-50 border border-slate-200/80 rounded-2xl p-4 flex flex-col gap-3">
                <div className="flex items-center gap-2">
                  <div className="bg-indigo-100 p-1.5 rounded-lg text-indigo-600">
                    <School style={{ width: '18px', height: '18px' }} />
                  </div>
                  <h3 className="font-extrabold text-xs uppercase tracking-wider text-slate-800">Mi Aula: Secciones y Alumnos</h3>
                </div>
                <p className="text-xs text-slate-500 font-medium leading-relaxed">
                  ¿Quieres guardar tus salones y alumnos? Accede gratis con Google de forma automática para gestionar tus listas de estudiantes hoy mismo.
                </p>
                <Button 
                  onClick={handleGoogleSignIn}
                  variant="outline"
                  className="w-full cursor-pointer hover:bg-slate-100 font-extrabold text-xs uppercase border-slate-200 text-slate-700 py-5 rounded-xl shadow-sm flex items-center justify-center gap-2"
                >
                  <svg className="w-4 h-4 mr-1" viewBox="0 0 24 24">
                    <path fill="#EA4335" d="M12.24 10.285V13.4h6.86c-.277 1.56-1.602 4.585-6.86 4.585-4.54 0-8.24-3.765-8.24-8.4s3.7-8.4 8.24-8.4c2.58 0 4.307 1.095 5.298 2.045l2.465-2.37C18.555 1.21 15.625 0 12.24 0 5.58 0 0 5.37 0 12s5.58 12 12.24 12c6.96 0 11.57-4.83 11.57-11.79 0-.795-.085-1.4-.185-1.925H12.24z"/>
                  </svg>
                  ACCEDER CON GOOGLE
                </Button>
              </div>
            ) : (
              <div className="bg-indigo-50/40 border border-indigo-100 rounded-2xl p-4 flex flex-col gap-4">
                {/* Header Auth User */}
                <div className="flex items-center justify-between">
                  <div className="flex items-center gap-2">
                    {currentUser.photoURL ? (
                      <img src={currentUser.photoURL} className="w-7 h-7 rounded-full border border-indigo-200" referrerPolicy="no-referrer" />
                    ) : (
                      <div className="w-7 h-7 rounded-full bg-indigo-200 text-indigo-700 font-bold flex items-center justify-center text-xs">
                        {String(currentUser.displayName || "D").charAt(0).toUpperCase()}
                      </div>
                    )}
                    <div className="flex flex-col">
                      <span className="font-extrabold text-xs text-slate-800 leading-none">Hola, {currentUser.displayName || "Docente"}</span>
                      <span className="text-[10px] text-indigo-600 font-medium">{currentUser.email}</span>
                    </div>
                  </div>
                  <button 
                    onClick={handleSignOut}
                    className="text-slate-400 hover:text-red-550 transition-colors flex items-center gap-1 text-[10px] uppercase font-extrabold cursor-pointer"
                    title="Cerrar sesión"
                  >
                    <LogOut style={{ width: '12px', height: '12px' }} /> Salir
                  </button>
                </div>

                {/* Secciones Dashboard Action Blocks */}
                <div className="flex flex-col gap-3">
                  <div className="flex items-center justify-between border-t border-indigo-100/60 pt-3">
                    <span className="font-extrabold text-xs text-indigo-900 uppercase tracking-wider flex items-center gap-1.5">
                      <Layers style={{ width: '14px', height: '14px' }} className="text-indigo-500" /> Mis Secciones ({secciones.length})
                    </span>
                    <Button 
                      onClick={() => {
                        setEditingSection(null);
                        setSectionGrado("");
                        setSectionNivel("");
                        setSectionNombre("");
                        setSectionRawStudentsText("");
                        setShowSectionModal(true);
                      }}
                      size="sm"
                      className="bg-indigo-600 hover:bg-indigo-700 text-white font-bold uppercase tracking-wider text-[9px] rounded-lg h-7 px-2.5 gap-1 shadow-sm cursor-pointer"
                    >
                      <Plus style={{ width: '12px', height: '12px' }} /> NUEVA SECCIÓN
                    </Button>
                  </div>

                  {loadingSecciones ? (
                    <div className="text-center py-4 text-xs font-semibold text-indigo-400 animate-pulse">Sincronizando con la nube...</div>
                  ) : secciones.length === 0 ? (
                    <div className="text-center py-5 border-2 border-dashed border-indigo-200/50 rounded-xl bg-indigo-50/20">
                      <p className="text-xs text-slate-400 font-bold uppercase">No tienes secciones registradas</p>
                      <p className="text-[10px] text-slate-400 mt-1">Registra tu primer salón de clase para empezar.</p>
                    </div>
                  ) : (
                    <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 max-h-[220px] overflow-y-auto no-scrollbar">
                      {secciones.map((sec) => {
                        const styleConfig = getColorClasses(sec.color);
                        return (
                          <div 
                            key={sec.id} 
                            className={`border rounded-xl shadow-sm flex items-center justify-between transition-all p-2.5 ${styleConfig.border} ${styleConfig.bgLight}`}
                          >
                            <div className="flex flex-col min-w-0 pr-2 select-none">
                              <span 
                                className={`font-black text-sm leading-tight truncate uppercase ${styleConfig.text}`}
                                title={`${sec.grado} - ${sec.nombre}`}
                              >
                                {sec.nombre}
                              </span>
                            </div>
                            <div className="flex items-center gap-1 shrink-0">
                              <button 
                                onClick={() => {
                                  setEditingSection(sec);
                                  setSectionGrado(sec.grado);
                                  setSectionNivel(sec.nivel);
                                  setSectionNombre(sec.nombre);
                                  setSectionColor(sec.color || "indigo");
                                  setSectionRawStudentsText("");
                                  setShowSectionModal(true);
                                }}
                                className="bg-slate-50 hover:bg-slate-100 text-slate-600 p-1.5 rounded-lg border border-slate-200/80 transition-colors cursor-pointer flex items-center"
                                title="Editar Aula o Color"
                              >
                                <Pencil style={{ width: '11px', height: '11px' }} />
                              </button>
                              <button 
                                onClick={() => openManageStudents(sec)}
                                className="bg-slate-50 hover:bg-slate-100 text-slate-600 p-1.5 rounded-lg border border-slate-200/80 transition-colors cursor-pointer flex items-center"
                                title="Gestionar Alumnos (Ingresar / Eliminar)"
                              >
                                <Users style={{ width: '11px', height: '11px' }} />
                              </button>
                              <button 
                                onClick={() => handleDeleteSection(sec.id)}
                                className="bg-rose-50 hover:bg-rose-100 text-rose-600 p-1.5 rounded-lg border border-rose-100 transition-colors cursor-pointer flex items-center"
                                title="Eliminar Sección"
                              >
                                <Trash2 style={{ width: '11px', height: '11px' }} />
                              </button>
                              <Button
                                onClick={() => handleStartGameWithSection(sec)}
                                size="sm"
                                className="bg-emerald-500 hover:bg-emerald-600 hover:scale-[1.08] active:scale-95 text-white font-extrabold rounded-lg h-[28px] w-[28px] p-0 cursor-pointer transition-all shrink-0 flex items-center justify-center shadow-sm"
                                title="Aprender hoy"
                              >
                                <Play style={{ width: '10px', height: '10px' }} className="fill-current" />
                              </Button>
                            </div>
                          </div>
                        );
                      })}
                    </div>
                  )}
                </div>
              </div>
            )}
          </div>

          <div 
            className="grid grid-cols-3 border-t border-slate-200/60 w-full pt-4 md:pt-6" 
            style={{ 
              gap: 'var(--fluid-space-xs)', 
              maxWidth: 'min(30rem, 100%)'
            }}
          >
            <div className="flex flex-col items-center md:items-start" style={{ gap: '4px' }}>
              <div className="bg-amber-400/10 rounded-xl border border-amber-400/20 p-2 shadow-sm">
                <Sparkles className="text-amber-500" style={{ width: '14px', height: '14px' }} />
              </div>
              <span className="font-extrabold uppercase tracking-wider text-slate-400" style={{ fontSize: '9px' }}>IA Socrática</span>
            </div>
            <div className="flex flex-col items-center md:items-start" style={{ gap: '4px' }}>
              <div className="bg-blue-400/10 rounded-xl border border-blue-400/20 p-2 shadow-sm">
                <Users className="text-blue-500" style={{ width: '14px', height: '14px' }} />
              </div>
              <span className="font-extrabold uppercase tracking-wider text-slate-400" style={{ fontSize: '9px' }}>Cooperativo</span>
            </div>
            <div className="flex flex-col items-center md:items-start" style={{ gap: '4px' }}>
              <div className="bg-emerald-400/10 rounded-xl border border-emerald-400/20 p-2 shadow-sm">
                <Upload className="text-emerald-500" style={{ width: '14px', height: '14px' }} />
              </div>
              <span className="font-extrabold uppercase tracking-wider text-slate-400" style={{ fontSize: '9px' }}>Exportar</span>
            </div>
          </div>
        </div>

        {/* Right Side: Youth Social Group Game Polaroid Representation with sliding photo gallery */}
        <div className="md:col-span-6 flex items-center justify-center w-full">
          <motion.div 
            initial={{ opacity: 0, scale: 0.9, rotate: -2 }}
            animate={{ 
              opacity: 1, 
              scale: 1, 
              rotate: 1,
              y: [-6, 6, -6],
            }}
            transition={{
              y: {
                duration: 5,
                repeat: Infinity,
                ease: "easeInOut"
              },
              default: { duration: 0.8 }
            }}
            className="bg-white p-[var(--fluid-space-sm)] rounded-[2.5rem] shadow-[0_25px_60px_rgba(30,41,59,0.18)] border-2 border-slate-100/95 w-full max-w-[28rem] relative group"
          >
            {/* Polaroid Gloss Overlay/Gloss Reflection Effect */}
            <div className="absolute inset-0 bg-gradient-to-tr from-transparent via-white/10 to-white/15 pointer-events-none rounded-[2.5rem] z-20" />
            
            {/* The beautiful image container with sliding transitions */}
            <div className="aspect-[4/3] rounded-2xl overflow-hidden relative border border-slate-200/80 ring-2 ring-indigo-50/50 bg-slate-100 shrink-0 shadow-inner select-none">
              <AnimatePresence mode="wait">
                <motion.img 
                  key={startImageIdx}
                  src={START_SLIDES[startImageIdx].src} 
                  alt={START_SLIDES[startImageIdx].title}
                  initial={{ opacity: 0, x: 25, scale: 1.03 }}
                  animate={{ opacity: 1, x: 0, scale: 1 }}
                  exit={{ opacity: 0, x: -25, scale: 0.97 }}
                  transition={{ duration: 0.45, ease: "easeInOut" }}
                  className="w-full h-full object-cover"
                  referrerPolicy="no-referrer"
                />
              </AnimatePresence>

              {/* Glowing Dynamic virtual bubble callout badge */}
              <div className="absolute top-3 left-3 bg-indigo-600/90 hover:bg-indigo-600 backdrop-blur-md text-white font-black text-[9px] uppercase tracking-widest px-3 py-1.5 rounded-full border border-white/20 flex items-center gap-1.5 shadow-md">
                <Users style={{ width: '10px', height: '10px' }} /> {START_SLIDES[startImageIdx].bubble}
              </div>

              {/* Slider Pagination Controls - Left & Right Dots / Controls */}
              <div className="absolute bottom-3 right-3 flex items-center gap-1 bg-black/40 backdrop-blur-md p-1 px-2.5 rounded-full z-10 border border-white/10">
                <button 
                  onClick={(e) => {
                    e.stopPropagation();
                    setStartImageIdx(prev => (prev - 1 + START_SLIDES.length) % START_SLIDES.length);
                  }}
                  className="text-white/80 hover:text-white transition-colors cursor-pointer"
                  title="Anterior"
                >
                  <ChevronLeft style={{ width: '14px', height: '14px' }} />
                </button>
                <div className="flex gap-1 px-1">
                  {START_SLIDES.map((_, idx) => (
                    <div 
                      key={idx} 
                      className={`h-1.5 rounded-full transition-all duration-300 ${startImageIdx === idx ? 'w-3 bg-indigo-400' : 'w-1.5 bg-white/40'}`} 
                    />
                  ))}
                </div>
                <button 
                  onClick={(e) => {
                    e.stopPropagation();
                    setStartImageIdx(prev => (prev + 1) % START_SLIDES.length);
                  }}
                  className="text-white/80 hover:text-white transition-colors cursor-pointer"
                  title="Siguiente"
                >
                  <ChevronRight style={{ width: '14px', height: '14px' }} />
                </button>
              </div>
            </div>

            {/* Dynamic, responsive details explaining game context */}
            <div className="pt-4 pb-2 text-center flex flex-col items-center">
              <AnimatePresence mode="wait">
                <motion.div
                  key={startImageIdx}
                  initial={{ opacity: 0, y: 5 }}
                  animate={{ opacity: 1, y: 0 }}
                  exit={{ opacity: 0, y: -5 }}
                  transition={{ duration: 0.3 }}
                  className="flex flex-col items-center w-full px-1"
                >
                  <h3 className="text-slate-800 font-extrabold text-sm uppercase tracking-wide leading-none mb-1">
                    {START_SLIDES[startImageIdx].title}
                  </h3>
                  <p className="text-[11px] text-slate-500 font-medium leading-relaxed max-w-[24rem]">
                    {START_SLIDES[startImageIdx].desc}
                  </p>
                  <p className="text-[10px] text-indigo-600 font-black uppercase mt-1.5 tracking-wider">
                    {START_SLIDES[startImageIdx].hashtag}
                  </p>
                </motion.div>
              </AnimatePresence>
            </div>
            
            {/* Extra Visual floating badges and stickers */}
            <div className="absolute -top-3 -left-3 bg-pink-500 text-white font-black text-[10px] tracking-wider uppercase px-4 py-1.5 rounded-2xl border-2 border-white shadow-lg rotate-[-8deg] animate-pulse">
              🎮 ¡Dinamismo Extremo!
            </div>
            <div className="absolute -bottom-3 -right-3 bg-amber-500 text-white font-black text-[10px] tracking-wider uppercase px-4 py-1.5 rounded-2xl border-2 border-white shadow-lg rotate-[6deg]">
              ✨ ¡Súper Cooperativo!
            </div>
          </motion.div>
        </div>

      </motion.div>
    );
  };

  const renderConfig = () => (
    <motion.div 
      initial={{ opacity: 0, scale: 0.98 }}
      animate={{ opacity: 1, scale: 1 }}
      className="w-full h-full flex flex-col mx-auto overflow-hidden bg-slate-50/50"
      style={{ 
        maxWidth: 'min(1850px, 100%)', 
        padding: 'var(--fluid-space-sm)',
        gap: 'var(--fluid-space-xs)'
      }}
    >
      <div className="flex items-center justify-between shrink-0" style={{ marginBottom: 'var(--fluid-space-xs)' }}>
        <div className="flex items-center">
          <Button 
            variant="ghost" 
            size="icon" 
            onClick={() => setScreen("start")} 
            className="mr-3 text-slate-400 hover:bg-white p-2 rounded-xl shadow-md border border-slate-100 transition-all hover:scale-110 active:scale-90"
            style={{ width: 'var(--fluid-space-md)', height: 'var(--fluid-space-md)' }}
          >
            <ChevronLeft style={{ width: 'var(--fluid-space-xs)', height: 'var(--fluid-space-xs)' }} />
          </Button>
          <div className="flex flex-col">
            <h2 className="font-black text-slate-900 italic tracking-tighter uppercase leading-tight py-1" style={{ fontSize: 'var(--fluid-text-2xl)' }}>
              Configuración
            </h2>
          </div>

          <div className="flex bg-slate-100 p-1 rounded-2xl border border-slate-200 ml-6 shrink-0">
            <button
              onClick={() => { setConfigTab("general"); setIsExcelImportActive(false); }}
              className={`px-4 py-1.5 rounded-xl font-extrabold text-xs uppercase tracking-wider transition-all cursor-pointer ${configTab === 'general' ? 'bg-white text-slate-900 shadow-sm' : 'text-slate-500 hover:text-slate-700'}`}
            >
              ⚙️ Reglas y Carga
            </button>
            <button
              onClick={() => {
                if (loadedCards.length === 0) {
                  setLoadedCards(DEFAULT_CARDS.map((c, i) => ({
                    ...c,
                    id: `default-${i}-${Date.now()}`
                  })));
                }
                setConfigTab("deck");
                setIsExcelImportActive(false);
              }}
              className={`px-4 py-1.5 rounded-xl font-extrabold text-xs uppercase tracking-wider transition-all cursor-pointer flex items-center gap-1.5 ${configTab === 'deck' ? 'bg-white text-slate-900 shadow-sm' : 'text-slate-500 hover:text-slate-700'}`}
            >
              🗂️ Administrador de Baraja
              <Badge className="bg-blue-600 text-white font-black text-[10px] px-1.5 py-0 h-5" variant="secondary">
                {loadedCards.length > 0 ? loadedCards.length : DEFAULT_CARDS.length}
              </Badge>
            </button>
          </div>
        </div>
        
        <div className="hidden sm:flex items-center">
          <div className="flex items-center gap-2 px-4 py-1.5 bg-white rounded-xl border border-slate-100 shadow-sm">
            <div className="w-2 h-2 rounded-full bg-emerald-500 animate-pulse" />
            <span className="font-black text-slate-600 uppercase tracking-tighter" style={{ fontSize: 'var(--fluid-text-xs)' }}>
              {getAvailableCards().length} Tarjetas activas
            </span>
          </div>
        </div>
      </div>

      {configTab === "deck" ? (
        <div className="flex-1 flex flex-col min-h-0 bg-slate-50/20 p-6 rounded-[2rem] border border-slate-100/50 overflow-y-auto hidden-scrollbar gap-6">
          
          {/* Bento Group: Paso 1 and Paso 3 side by side */}
          <div className="grid grid-cols-1 lg:grid-cols-12 gap-6 shrink-0">
            
            {/* Paso 1 Card */}
            <div className="lg:col-span-8 bg-white rounded-[24px] p-5 border border-slate-100 shadow-sm flex flex-col gap-4">
              <ExcelImporter 
                loadedCardsLength={(loadedCards.length > 0 ? loadedCards : DEFAULT_CARDS).length}
                onImportComplete={(newCards) => {
                  const normalized = newCards.map((c, i) => ({
                    ...c,
                    id: c.id || `uploaded-${i}-${Date.now()}`
                  }));
                  setLoadedCards(normalized);
                }}
                onResetSet={() => {
                  setLoadedCards(DEFAULT_CARDS.map((c, i) => ({
                    ...c,
                    id: `default-${i}-${Date.now()}`
                  })));
                  toast.success("¡Set de ejemplo (8 cartas) cargado con éxito!");
                }}
                onClearSet={() => {
                  setLoadedCards([]);
                  toast.success("Baraja vaciada. Puedes importar o crear tarjetas manualmente.");
                }}
              />
            </div>

            {/* Paso 3 Card */}
            <div className="lg:col-span-4 bg-white rounded-[24px] p-5 border border-slate-100 shadow-sm flex flex-col justify-between gap-4">
              <div className="flex items-center gap-2.5 border-b border-slate-100 pb-3">
                <div className="w-1.5 h-6 bg-blue-600 rounded-full shrink-0" />
                <h2 className="text-slate-900 font-extrabold text-sm uppercase tracking-wide">
                  PASO 3: DESCARGAR WORD
                </h2>
              </div>
              
              <p className="text-[11px] leading-relaxed text-slate-400 font-bold uppercase tracking-wide">
                Configura el nombre del documento impresor y la distribución de las columnas simétricas. Las celdas se adaptarán para crear una baraja perfecta.
              </p>

              <div className="space-y-3">
                {/* Filename Input */}
                <div className="space-y-1">
                  <label className="text-[10px] font-black text-slate-500 uppercase tracking-widest pl-0.5">Nombre del archivo descargable:</label>
                  <div className="flex bg-white border border-slate-200 rounded-xl overflow-hidden focus-within:ring-2 focus-within:ring-blue-500 transition-all h-11 items-center px-4 gap-2">
                    <input 
                      type="text" 
                      value={downloadFileName} 
                      onChange={(e) => setDownloadFileName(e.target.value)} 
                      className="flex-1 bg-transparent border-none outline-none text-slate-800 font-extrabold text-xs" 
                      placeholder="tarjetas_tabu_editables"
                    />
                    <div className="h-full flex items-center border-l border-slate-200 pl-3">
                      <span className="text-slate-400 font-black text-[10px] uppercase tracking-wider">.docx</span>
                    </div>
                  </div>
                </div>

                {/* Print Column Formats Selector */}
                <div className="space-y-1">
                  <label className="text-[10px] font-black text-slate-500 uppercase tracking-widest pl-0.5">Formato de columnas en el impresor:</label>
                  <div className="grid grid-cols-3 gap-2">
                    {[2, 3, 4].map((cols) => (
                      <button
                        key={cols}
                        type="button"
                        onClick={() => setExportColumns(cols)}
                        className={`h-11 font-bold text-xs uppercase rounded-xl border transition-all flex flex-col items-center justify-center ${
                          exportColumns === cols
                            ? 'border-blue-600 bg-white text-blue-600 font-extrabold shadow-sm'
                            : 'border-slate-200 bg-white text-slate-500 hover:border-slate-350 hover:text-slate-800'
                        }`}
                      >
                        {cols} Cols
                      </button>
                    ))}
                  </div>
                  <div className="flex justify-center pt-1">
                    {exportColumns === 3 ? (
                      <span className="text-[10px] font-medium text-slate-400 italic">
                        * Tamaño de naipes ideal
                      </span>
                    ) : (
                      <span className="text-[8px] font-bold text-slate-400 uppercase tracking-widest pl-1">Máximo por página, idénticas proporciones</span>
                    )}
                  </div>
                </div>
              </div>

              {/* Separator and CTA Export Word Button */}
              <div className="border-t border-slate-100 pt-3">
                <Button 
                  onClick={() => {
                    const cardsToExport = loadedCards.length > 0 ? loadedCards : DEFAULT_CARDS.map((c, i) => ({ ...c, id: `default-${i}` }));
                    exportDeckToWord(cardsToExport, exportColumns, downloadFileName);
                    toast.success("¡Documento Word exportado con éxito!");
                  }}
                  className="w-full bg-[#4f46e5] hover:bg-[#4338ca] text-white font-extrabold rounded-xl text-xs uppercase tracking-wider h-11 transition-all shadow-md flex items-center justify-center gap-2"
                >
                  <Download size={14} /> Descargar Word editable (.docx)
                </Button>
                <p className="text-[9px] text-slate-405 font-bold uppercase tracking-wider text-center select-none mt-2">
                  Descarga local instantánea 100% segura y privada
                </p>
              </div>
            </div>

          </div>

          {/* Paso 2: Card Grid (fully self-scrolling and self-contained inside CardGrid.tsx) */}
          <div className="flex-1 min-h-0">
            <CardGrid 
              cards={loadedCards.length > 0 ? loadedCards : DEFAULT_CARDS.map((c, i) => ({ ...c, id: `default-${i}-${Date.now()}` }))} 
              onCardsChange={setLoadedCards} 
            />
          </div>

          {/* Launch App Play Screen Button at Footer */}
          <div className="pt-2 flex justify-end shrink-0">
            <Button 
              size="lg" 
              className="rounded-xl bg-blue-600 hover:bg-blue-700 text-white font-black uppercase tracking-[0.2em] shadow-lg active:scale-95 transition-all h-12 px-8 text-xs shrink-0"
              onClick={proceedToTeams}
            >
              EMPEZAR JUEGO CON ESTA BARAJA <ChevronLeft className="rotate-180 w-5 h-5 ml-2" />
            </Button>
          </div>
        </div>
      ) : (
        <div className="grid lg:grid-cols-3 items-stretch flex-1 min-h-0 w-full overflow-hidden" style={{ gap: 'var(--fluid-space-md)' }}>
        
        {/* Column 1: Game Rules */}
        <motion.div 
          initial={{ opacity: 0, x: -20 }}
          animate={{ opacity: 1, x: 0 }}
          transition={{ delay: 0.05 }}
          className="flex flex-col h-full"
        >
          <Card className="glass-panel border-none shadow-xl rounded-[var(--radius-2xl)] bg-white flex flex-col flex-1 hover:shadow-blue-500/5 transition-all duration-300">
            <CardHeader className="pb-1.5 pt-3 px-4">
              <div className="flex items-center gap-3">
                <div className="bg-blue-600 text-white p-1.5 rounded-xl shadow-lg">
                   <Settings size={16} />
                </div>
                <CardTitle className="text-slate-900 font-black uppercase tracking-tight" style={{ fontSize: 'var(--fluid-text-sm)' }}>
                  Reglas
                </CardTitle>
              </div>
            </CardHeader>
            <CardContent className="px-4 pb-3 space-y-2 flex-1 overflow-y-auto hidden-scrollbar">
              {/* Teams */}
              <div className="space-y-2">
                <div className="flex justify-between items-center">
                  <Label className="font-black text-slate-800 uppercase tracking-widest flex items-center gap-2" style={{ fontSize: 'var(--fluid-text-sm)' }}>
                    <Users size={16} className="text-blue-500" /> Equipos
                  </Label>
                  <Badge className="bg-blue-600 text-white font-black px-2 py-0.5 rounded-lg" style={{ fontSize: 'var(--fluid-text-base)' }}>{numTeams}</Badge>
                </div>
                <div className="flex items-center gap-3 py-1">
                  <Button variant="outline" size="icon" className="rounded-xl border-slate-200 h-8 w-8 shrink-0 font-black" onClick={() => safeUpdate(setNumTeams, numTeams - 1, 2, 6)}><Minus size={14} /></Button>
                  <Slider value={[numTeams]} min={2} max={6} step={1} onValueChange={(v) => safeUpdate(setNumTeams, v[0], 2, 6)} className="flex-1" />
                  <Button variant="outline" size="icon" className="rounded-xl border-slate-200 h-8 w-8 shrink-0 font-black" onClick={() => safeUpdate(setNumTeams, numTeams + 1, 2, 6)}><Plus size={14} /></Button>
                </div>
              </div>

              {/* Rounds */}
              <div className="space-y-2">
                <div className="flex justify-between items-center">
                  <Label className="font-black text-slate-800 uppercase tracking-widest flex items-center gap-2" style={{ fontSize: 'var(--fluid-text-sm)' }}>
                    <RotateCcw size={16} className="text-blue-500" /> Rondas
                  </Label>
                  <Badge className="bg-blue-600 text-white font-black px-2 py-0.5 rounded-lg" style={{ fontSize: 'var(--fluid-text-base)' }}>{totalRounds}</Badge>
                </div>
                <div className="flex items-center gap-3 py-1">
                  <Button variant="outline" size="icon" className="rounded-xl border-slate-200 h-8 w-8 shrink-0 font-black" onClick={() => safeUpdate(setTotalRounds, totalRounds - 1, 1, 10)}><Minus size={14} /></Button>
                  <Slider value={[totalRounds]} min={1} max={10} step={1} onValueChange={(v) => safeUpdate(setTotalRounds, v[0], 1, 10)} className="flex-1" />
                  <Button variant="outline" size="icon" className="rounded-xl border-slate-200 h-8 w-8 shrink-0 font-black" onClick={() => safeUpdate(setTotalRounds, totalRounds + 1, 1, 10)}><Plus size={14} /></Button>
                </div>
              </div>

              {/* Time */}
              <div className="space-y-2">
                <div className="flex justify-between items-center">
                  <Label className="font-black text-slate-800 uppercase tracking-widest flex items-center gap-2" style={{ fontSize: 'var(--fluid-text-sm)' }}>
                    <Timer size={16} className="text-blue-500" /> Tiempo
                  </Label>
                  <Badge className="bg-blue-600 text-white font-black px-2 py-0.5 rounded-lg" style={{ fontSize: 'var(--fluid-text-base)' }}>{timePerTurn}s</Badge>
                </div>
                <div className="flex items-center gap-3 py-1">
                  <Button variant="outline" size="icon" className="rounded-xl border-slate-200 h-8 w-8 shrink-0 font-black" onClick={() => safeUpdate(setTimePerTurn, timePerTurn - 10, 10, 120)}><Minus size={14} /></Button>
                  <Slider value={[timePerTurn]} min={10} max={120} step={10} onValueChange={(v) => safeUpdate(setTimePerTurn, v[0], 10, 120)} className="flex-1" />
                  <Button variant="outline" size="icon" className="rounded-xl border-slate-200 h-8 w-8 shrink-0 font-black" onClick={() => safeUpdate(setTimePerTurn, timePerTurn + 10, 10, 120)}><Plus size={14} /></Button>
                </div>
              </div>

              {/* Passes */}
              <div className="space-y-2">
                <div className="flex justify-between items-center">
                  <Label className="font-black text-slate-800 uppercase tracking-widest flex items-center gap-2" style={{ fontSize: 'var(--fluid-text-sm)' }}>
                    <SkipForward size={16} className="text-blue-500" /> Pases
                  </Label>
                  <Badge className="bg-blue-600 text-white font-black px-2 py-0.5 rounded-lg" style={{ fontSize: 'var(--fluid-text-base)' }}>{passesPerTurn}</Badge>
                </div>
                <div className="flex items-center gap-3 py-1">
                  <Button variant="outline" size="icon" className="rounded-xl border-slate-200 h-8 w-8 shrink-0 font-black" onClick={() => safeUpdate(setPassesPerTurn, passesPerTurn - 1, 0, 10)}><Minus size={14} /></Button>
                  <Slider value={[passesPerTurn]} min={0} max={10} step={1} onValueChange={(v) => safeUpdate(setPassesPerTurn, v[0], 0, 10)} className="flex-1" />
                  <Button variant="outline" size="icon" className="rounded-xl border-slate-200 h-8 w-8 shrink-0 font-black" onClick={() => safeUpdate(setPassesPerTurn, passesPerTurn + 1, 0, 10)}><Plus size={14} /></Button>
                </div>
              </div>
            </CardContent>
          </Card>
        </motion.div>

        {/* Column 2: Content Selection & AI */}
        <motion.div 
          initial={{ opacity: 0, y: 20 }}
          animate={{ opacity: 1, y: 0 }}
          transition={{ delay: 0.1 }}
          className="space-y-4 flex flex-col h-full"
        >
          <Card className="glass-panel border-none shadow-xl rounded-[var(--radius-2xl)] bg-white overflow-hidden hover:shadow-amber-500/5 transition-all duration-300">
            <CardHeader className="pb-1.5 pt-3 px-4">
              <div className="flex items-center gap-3">
                <div className="bg-amber-500 text-white p-1.5 rounded-xl shadow-lg">
                   <Sparkles size={16} />
                </div>
                <CardTitle className="text-slate-900 font-black uppercase tracking-tight" style={{ fontSize: 'var(--fluid-text-sm)' }}>
                  Contenido
                </CardTitle>
              </div>
            </CardHeader>
            <CardContent className="px-4 pb-3 space-y-2">
              <div className="space-y-2">
                <Label className="font-black text-slate-800 uppercase tracking-widest pl-1" style={{ fontSize: 'var(--fluid-text-sm)' }}>Generar con IA</Label>
                <div className="flex gap-2">
                  <Input 
                    value={aiTopic}
                    onChange={(e) => setAiTopic(e.target.value)}
                    placeholder="Ej: Mitología..."
                    className="bg-slate-50 border-slate-200 text-slate-900 rounded-xl focus:ring-amber-500 font-black h-10 px-4 shadow-inner"
                    style={{ fontSize: 'var(--fluid-text-sm)' }}
                  />
                  <Button 
                    variant="secondary"
                    onClick={generateWithAI} 
                    className="rounded-xl bg-amber-500 hover:bg-amber-600 text-white font-black h-10 shadow-lg transition-all hover:scale-105 active:scale-95 px-5"
                    disabled={isGenerating || !aiTopic}
                  >
                    {isGenerating ? <RotateCcw className="animate-spin w-4 h-4" /> : "IR"}
                  </Button>
                </div>
              </div>

              <div className="space-y-2 pt-2 border-t border-slate-50">
                <Label className="font-black text-slate-800 uppercase tracking-widest pl-1" style={{ fontSize: 'var(--fluid-text-sm)' }}>Categoría</Label>
                <Select value={selectedCategory} onValueChange={setSelectedCategory}>
                  <SelectTrigger className="bg-slate-50 border-slate-200 text-slate-900 rounded-xl font-black h-10 px-4 shadow-inner text-left">
                    <SelectValue placeholder="Todas" />
                  </SelectTrigger>
                  <SelectContent className="rounded-xl border-slate-200 shadow-2xl">
                    <SelectItem value="all" className="font-black uppercase tracking-tight py-2 cursor-pointer">Todas</SelectItem>
                    {getCategories().map(cat => (
                      <SelectItem key={cat} value={cat} className="font-black uppercase tracking-tight py-2 cursor-pointer">{cat}</SelectItem>
                    ))}
                  </SelectContent>
                </Select>
              </div>

              <div className="space-y-2 pt-2 border-t border-slate-50">
                <Label className="font-black text-slate-800 uppercase tracking-widest pl-1" style={{ fontSize: 'var(--fluid-text-sm)' }}>Nivel de Complejidad (Logro)</Label>
                <Select value={selectedNivelLogro} onValueChange={setSelectedNivelLogro}>
                  <SelectTrigger className="bg-slate-50 border-slate-200 text-slate-900 rounded-xl font-black h-10 px-4 shadow-inner text-left">
                    <SelectValue placeholder="Todos los niveles" />
                  </SelectTrigger>
                  <SelectContent className="rounded-xl border-slate-200 shadow-2xl">
                    <SelectItem value="all" className="font-black uppercase tracking-tight py-2 cursor-pointer">Todos los Niveles</SelectItem>
                    <SelectItem value="Inicio" className="font-black uppercase tracking-tight py-2 cursor-pointer">Inicio (Fácil)</SelectItem>
                    <SelectItem value="Proceso" className="font-black uppercase tracking-tight py-2 cursor-pointer">Proceso (Medio)</SelectItem>
                    <SelectItem value="Logrado" className="font-black uppercase tracking-tight py-2 cursor-pointer">Logrado (Difícil)</SelectItem>
                    <SelectItem value="Destacado" className="font-black uppercase tracking-tight py-2 cursor-pointer font-bold">Destacado (Complejo)</SelectItem>
                  </SelectContent>
                </Select>
              </div>
            </CardContent>
          </Card>

          <Card className="glass-panel border-none shadow-xl rounded-[var(--radius-2xl)] bg-blue-600/5 border-2 border-dashed border-blue-100 flex flex-col justify-center items-center text-center p-3 h-32 mt-auto">
             <motion.div 
               animate={{ y: [0, -4, 0] }}
               transition={{ repeat: Infinity, duration: 4 }}
               className="bg-blue-600 text-white p-2 rounded-2xl shadow-xl mb-2"
             >
                <Users size={20} />
             </motion.div>
             <h3 className="font-black text-slate-900 uppercase tracking-tighter" style={{ fontSize: 'var(--fluid-text-sm)' }}>Equipos</h3>
             <p className="text-slate-500 font-black uppercase tracking-widest leading-tight mt-1" style={{ fontSize: '10px' }}>Personaliza nombres en el siguiente paso</p>
          </Card>
        </motion.div>

        {/* Column 3: Resources & Actions */}
        <motion.div 
          initial={{ opacity: 0, x: 20 }}
          animate={{ opacity: 1, x: 0 }}
          transition={{ delay: 0.15 }}
          className="space-y-4 flex flex-col h-full"
        >
           <Card className="glass-panel border-none shadow-xl rounded-[var(--radius-2xl)] bg-white overflow-hidden flex-1 flex flex-col min-h-0 hover:shadow-emerald-500/5 transition-all duration-300">
            <CardHeader className="pb-1.5 pt-3 px-4">
              <div className="flex items-center gap-3">
                <div className="bg-emerald-500 text-white p-1.5 rounded-xl shadow-lg">
                   <FileSpreadsheet size={16} />
                </div>
                <CardTitle className="text-slate-900 font-black uppercase tracking-tight" style={{ fontSize: 'var(--fluid-text-sm)' }}>
                  Recursos
                </CardTitle>
              </div>
            </CardHeader>
            <CardContent className="px-4 pb-3 space-y-2 flex-1 flex flex-col justify-center min-h-0 overflow-y-auto hidden-scrollbar">
               <button 
                 className="w-full h-20 rounded-2xl border-dashed border-2 border-slate-200 flex flex-col items-center justify-center gap-1 hover:bg-emerald-50 hover:border-emerald-300 hover:text-emerald-700 transition-all group overflow-hidden bg-slate-50 shrink-0"
                 onClick={() => setIsExcelImportActive(true)}
               >
                 <Upload size={24} className="text-slate-400 group-hover:scale-110 transition-transform" />
                 <span className="font-black uppercase tracking-widest" style={{ fontSize: '10px' }}>Cargar Excel Inteligente</span>
               </button>

               <div className="flex flex-col gap-2 pt-3 border-t border-slate-100 font-black shrink-0">
                  <Button 
                    variant="outline"
                    className="w-full justify-between rounded-xl bg-white border-slate-200 hover:bg-slate-900 hover:text-white transition-all h-10 px-4 font-black group"
                    onClick={() => downloadTemplate('cards')}
                  >
                    <span className="uppercase tracking-[0.1em]" style={{ fontSize: '10px' }}>Plantilla</span>
                    <Download size={16} className="group-hover:translate-y-1 transition-transform" />
                  </Button>
               </div>

               {loadedCards.length > 0 && (
                 <motion.div 
                   initial={{ opacity: 0, scale: 0.9 }}
                   animate={{ opacity: 1, scale: 1 }}
                   className="bg-emerald-50/50 rounded-2xl p-4 border border-emerald-100 flex items-center gap-4 shadow-inner shrink-0"
                 >
                   <div className="bg-emerald-500 p-2 rounded-xl shadow-md">
                      <CheckCircle2 className="text-white" size={18} />
                   </div>
                   <div className="flex flex-col">
                      <p className="font-black text-emerald-900 uppercase tracking-tight leading-none" style={{ fontSize: 'var(--fluid-text-xs)' }}>{loadedCards.length} Tarjetas</p>
                      <button onClick={() => setLoadedCards([])} className="text-[10px] font-black text-emerald-600 uppercase tracking-widest underline text-left mt-1.5 hover:text-emerald-800 transition-colors">Limpiar</button>
                   </div>
                 </motion.div>
               )}
            </CardContent>
          </Card>

          <Button 
            size="lg" 
            className="w-full rounded-[var(--radius-2xl)] bg-blue-600 hover:bg-blue-700 text-white font-black uppercase tracking-[0.4em] shadow-xl active:scale-95 transition-all group overflow-hidden h-16 relative shrink-0"
            onClick={proceedToTeams}
          >
            <div className="absolute inset-0 bg-gradient-to-r from-blue-400/0 via-white/10 to-blue-400/0 translate-x-[-100%] group-hover:translate-x-[100%] transition-transform duration-1000" />
            <span className="relative z-10 flex items-center justify-center gap-4" style={{ fontSize: 'var(--fluid-text-xl)' }}>
              CONTINUAR <ChevronLeft className="rotate-180 w-6 h-6 animate-pulse" />
            </span>
          </Button>
        </motion.div>
      </div>
      )}
    </motion.div>
  );

  const renderTeams = () => (
    <motion.div 
      initial={{ opacity: 0, x: 20 }}
      animate={{ opacity: 1, x: 0 }}
      className="w-full h-full flex flex-col mx-auto overflow-hidden"
      style={{ 
        maxWidth: 'min(1850px, 100%)', 
        padding: 'var(--fluid-space-md)',
        gap: 'var(--fluid-space-sm)'
      }}
    >
      <div className="flex items-center justify-between shrink-0" style={{ marginBottom: 'var(--fluid-space-sm)' }}>
        <div className="flex items-center">
          <Button 
            variant="ghost" 
            size="icon" 
            onClick={() => setScreen("config")} 
            className="mr-3 text-slate-400 hover:bg-slate-200/50 p-4 rounded-xl"
            style={{ width: 'var(--fluid-space-lg)', height: 'var(--fluid-space-lg)' }}
          >
            <ChevronLeft style={{ width: 'var(--fluid-space-sm)', height: 'var(--fluid-space-sm)' }} />
          </Button>
          <div className="flex flex-col">
            <h2 className="font-black text-slate-900 italic tracking-tighter uppercase leading-none" style={{ fontSize: 'var(--fluid-text-2xl)' }}>Group Maker</h2>
            <p className="font-bold text-slate-400 uppercase tracking-widest mt-1" style={{ fontSize: 'var(--fluid-text-xs)' }}>Organización de Integrantes</p>
          </div>
        </div>
        <div className="flex items-center" style={{ gap: 'var(--fluid-space-sm)' }}>
          <Button 
            variant="ghost" 
            onClick={() => downloadTemplate('students')}
            className="text-slate-400 hover:text-blue-500 font-black uppercase tracking-widest"
            style={{ fontSize: 'var(--fluid-text-xs)' }}
          >
            <Download className="mr-2" style={{ width: 'var(--fluid-space-xs)', height: 'var(--fluid-space-xs)' }} /> Plantilla
          </Button>
          <Button 
            variant="outline" 
            onClick={downloadCurrentTeams}
            disabled={teamMembers.slice(0, numTeams).every(t => (t?.length || 0) === 0)}
            className="border-slate-200 text-slate-600 font-gray-800 font-black uppercase tracking-widest rounded-xl hover:bg-slate-50 transition-colors"
            style={{ fontSize: 'var(--fluid-text-xs)', height: 'var(--fluid-space-xl)' }}
          >
            <Download className="mr-2" style={{ width: 'var(--fluid-space-xs)', height: 'var(--fluid-space-xs)' }} /> Descargar Equipos
          </Button>
          <div className="relative group">
            <Input type="file" accept=".xlsx, .xls" onChange={handleStudentsUpload} className="absolute inset-0 opacity-0 cursor-pointer z-10" />
            <Button variant="outline" className="border-slate-200 text-slate-600 font-black uppercase tracking-widest rounded-xl" style={{ fontSize: 'var(--fluid-text-xs)', height: 'var(--fluid-space-xl)' }}>
              <Upload className="mr-2" style={{ width: 'var(--fluid-space-xs)', height: 'var(--fluid-space-xs)' }} /> Excel
            </Button>
          </div>
        </div>
      </div>

      <div className="grid lg:grid-cols-12 flex-1 min-h-0 overflow-hidden w-full" style={{ gap: 'var(--fluid-space-sm)' }}>
        {/* Left Side: Input students */}
        <Card className="lg:col-span-3 glass-panel border-none shadow-xl rounded-[var(--radius-2xl)] flex flex-col bg-white/70 backdrop-blur-2xl overflow-hidden min-h-0">
          <CardHeader className="shrink-0" style={{ padding: 'var(--fluid-space-xs)' }}>
            <CardTitle className="text-slate-800 font-black uppercase tracking-widest flex items-center" style={{ gap: 'var(--fluid-space-xs)', fontSize: 'var(--fluid-text-base)' }}>
              <Users className="text-orange-500" style={{ width: 'var(--fluid-space-sm)', height: 'var(--fluid-space-sm)' }} /> Lista de Alumnos
            </CardTitle>
          </CardHeader>
          <CardContent className="flex-1 flex flex-col p-2 pt-0 min-h-0" style={{ gap: 'var(--fluid-space-xs)' }}>
            <textarea 
              value={studentList}
              onChange={(e) => setStudentList(e.target.value)}
              placeholder="Juan Pérez&#10;María Luz..."
              className="flex-1 w-full bg-slate-50/50 border border-slate-200 rounded-xl p-2 font-medium text-slate-700 focus:ring-2 focus:ring-orange-500 outline-none resize-none custom-scrollbar shadow-inner"
              style={{ fontSize: 'calc(var(--fluid-text-xs) * 0.9)' }}
            />
            <div className="shrink-0 flex flex-col" style={{ gap: 'var(--fluid-space-xs)', marginTop: 'var(--fluid-space-xs)' }}>
              <div className="grid grid-cols-2 gap-2 w-full">
                <Button 
                  onClick={shuffleTeams}
                  className="rounded-xl bg-orange-500 hover:bg-orange-600 text-white font-black uppercase tracking-widest shadow-lg shadow-orange-500/20 active:scale-95 transition-all w-full"
                  style={{ height: 'var(--fluid-space-lg)', fontSize: '10px' }}
                >
                  <RotateCcw className="mr-1.5" style={{ width: 'var(--fluid-space-xs)', height: 'var(--fluid-space-xs)' }} /> SORTEAR
                </Button>
                <Button 
                  variant="outline"
                  onClick={clearTeamsAndStudents}
                  className="rounded-xl border-rose-200 text-rose-600 hover:bg-rose-50 hover:text-rose-700 font-black uppercase tracking-widest active:scale-95 transition-all w-full"
                  style={{ height: 'var(--fluid-space-lg)', fontSize: '10px' }}
                >
                  <Trash2 className="mr-1.5" style={{ width: 'var(--fluid-space-xs)', height: 'var(--fluid-space-xs)' }} /> LIMPIAR
                </Button>
              </div>
              <Button 
                onClick={startGame}
                disabled={teamMembers.every(t => (t?.length || 0) === 0)}
                className="w-full rounded-xl bg-blue-600 hover:bg-blue-700 text-white font-black uppercase tracking-[0.2em] shadow-xl shadow-blue-500/20 active:scale-95 transition-all"
                style={{ height: 'var(--fluid-space-lg)', fontSize: 'var(--fluid-text-sm)' }}
              >
                JUGAR <Play className="ml-2 fill-current" style={{ width: 'var(--fluid-space-xs)', height: 'var(--fluid-space-xs)' }} />
              </Button>
            </div>
          </CardContent>
        </Card>

        {/* Right Side: Visual Groups */}
        <div className="lg:col-span-9 overflow-y-auto pr-1 custom-scrollbar grid grid-cols-1 md:grid-cols-2 xl:grid-cols-3 gap-3 content-start min-h-0 w-full">
          {Array.from({ length: numTeams }).map((_, tIdx) => (
            <motion.div 
              key={tIdx}
              layout
              initial={{ opacity: 0, y: 20 }}
              animate={{ opacity: 1, y: 0 }}
              transition={{ delay: tIdx * 0.1 }}
              onDragOver={handleDragOver}
              onDrop={() => handleDrop(tIdx)}
              className={`group h-fit rounded-[var(--radius-xl)] transition-all ${draggingMember && draggingMember.tIdx !== tIdx ? 'ring-2 ring-blue-500 ring-dashed bg-blue-50/30' : ''}`}
            >
              <Card className="overflow-hidden border-none shadow-xl rounded-[var(--radius-xl)] transition-all group-hover:scale-[1.01] bg-white">
                <div className={`${TEAM_COLORS[tIdx].bg} flex items-center justify-between`} style={{ padding: 'var(--fluid-space-xs) var(--fluid-space-sm)' }}>
                  <div className="flex items-center" style={{ gap: 'var(--fluid-space-xs)' }}>
                    <Users className="text-white opacity-40 shrink-0" style={{ width: 'var(--fluid-space-xs)', height: 'var(--fluid-space-xs)' }} />
                    <h3 className="text-white font-black uppercase tracking-tighter italic" style={{ fontSize: 'var(--fluid-text-sm)' }}>{TEAM_COLORS[tIdx].name}</h3>
                  </div>
                  <Badge variant="outline" className="text-white border-white/30 font-black px-1.5" style={{ fontSize: 'calc(var(--fluid-text-xs) * 0.8)' }}>
                    {teamMembers[tIdx]?.length || 0}
                  </Badge>
                </div>
                {/* Team settings bar inside card content */}
                <div className="bg-slate-50 border-b border-slate-100 px-3 py-1.5 flex items-center justify-between gap-1 shrink-0">
                  <span className="text-[10px] font-black uppercase text-slate-500 tracking-wider">Comunicadores:</span>
                  <div className="flex items-center gap-1 bg-white border border-slate-200 rounded-lg p-0.5 shadow-sm">
                    <button
                      onClick={() => {
                        const next = [...teamCommunicatorCount];
                        next[tIdx] = 1;
                        setTeamCommunicatorCount(next);
                      }}
                      className={`text-[9px] font-black px-1.5 py-0.5 rounded uppercase tracking-tighter transition-all ${
                        (teamCommunicatorCount[tIdx] || 1) === 1
                          ? 'bg-blue-600 text-white shadow-xs font-extrabold'
                          : 'text-slate-500 hover:bg-slate-50 font-bold'
                      }`}
                    >
                      1 solo
                    </button>
                    <button
                      onClick={() => {
                        const next = [...teamCommunicatorCount];
                        next[tIdx] = 2;
                        setTeamCommunicatorCount(next);
                      }}
                      className={`text-[9px] font-black px-1.5 py-0.5 rounded uppercase tracking-tighter transition-all ${
                        (teamCommunicatorCount[tIdx] || 1) === 2
                          ? 'bg-blue-600 text-white shadow-xs font-extrabold'
                          : 'text-slate-500 hover:bg-slate-50 font-bold'
                      }`}
                    >
                      De a 2
                    </button>
                  </div>
                </div>

                <CardContent className="p-0 min-h-[100px] max-h-[250px] overflow-y-auto custom-scrollbar">
                  <div className="divide-y divide-slate-50">
                    {(teamMembers[tIdx] || []).map((member, mIdx) => (
                      <motion.div 
                        key={member.id} 
                        initial={{ opacity: 0, x: -10 }}
                        animate={{ opacity: 1, x: 0 }}
                        draggable={true}
                        onDragStart={() => handleDragStart(tIdx, mIdx)}
                        className={`font-bold text-slate-700 flex items-center justify-between bg-white hover:bg-slate-50 transition-colors cursor-grab active:cursor-grabbing ${draggingMember?.tIdx === tIdx && draggingMember?.mIdx === mIdx ? 'opacity-40' : ''}`}
                        style={{ padding: '4px var(--fluid-space-sm)', fontSize: 'var(--fluid-text-sm)', lineHeight: '1.2' }}
                      >
                        <div className="flex items-center gap-1.5 truncate min-w-0 flex-1">
                          <button 
                            onClick={(e) => {
                              e.stopPropagation();
                              toggleMemberLock(tIdx, mIdx);
                            }}
                            className={`shrink-0 transition-colors ${member.locked ? 'text-amber-500' : 'text-slate-200 hover:text-slate-400'}`}
                            title={member.locked ? "Desbloquear" : "Bloquear en este equipo"}
                          >
                            {member.locked ? <Lock size={12} strokeWidth={3} /> : <LockOpen size={12} />}
                          </button>

                          {/* Role Toggle Button */}
                          <button
                            onClick={(e) => {
                              e.stopPropagation();
                              toggleMemberRole(tIdx, mIdx);
                            }}
                            className={`shrink-0 flex items-center justify-center p-0.5 rounded-md transition-all ${
                              (member.role || 'comunicador') === 'comunicador'
                                ? 'bg-emerald-50 text-emerald-600 hover:bg-emerald-100 border border-emerald-100'
                                : 'bg-orange-50 text-orange-600 hover:bg-orange-100 border border-orange-100'
                            }`}
                            title={(member.role || 'comunicador') === 'comunicador' ? 'Rol: Comunicador (Da pistas). Haz clic para cambiar a Adivinador.' : 'Rol: Adivinador (Adivina palabras). Haz clic para cambiar a Comunicador.'}
                            style={{ width: '18px', height: '18px' }}
                          >
                            {(member.role || 'comunicador') === 'comunicador' ? <Volume2 size={10} strokeWidth={3} /> : <Target size={10} strokeWidth={3} />}
                          </button>
                          
                          <span className="truncate uppercase tracking-tight font-black" style={{ fontSize: '11px' }}>{member.name}</span>
                        </div>
                        <div className="flex items-center shrink-0 gap-1.5">
                          <span className={`text-[8px] font-black uppercase px-1.5 py-0.5 rounded tracking-wider ${
                            (member.role || 'comunicador') === 'comunicador'
                              ? 'bg-emerald-100/80 text-emerald-800 border border-emerald-200/30'
                              : 'bg-orange-100/80 text-orange-800 border border-orange-200/30'
                          }`}>
                            {(member.role || 'comunicador') === 'comunicador' ? '📢 pistas' : '🎯 adivina'}
                          </span>

                          <button 
                            onClick={(e) => {
                              e.stopPropagation();
                              removeMember(tIdx, mIdx);
                            }}
                            className="text-slate-200 hover:text-rose-500 transition-colors p-0.5"
                          >
                            <X style={{ width: 'var(--fluid-space-xs)', height: 'var(--fluid-space-xs)' }} />
                          </button>
                        </div>
                      </motion.div>
                    ))}
                    {(teamMembers[tIdx]?.length || 0) === 0 && (
                      <div className="text-center opacity-30 flex flex-col items-center justify-center" style={{ padding: 'var(--fluid-space-md)', gap: '4px' }}>
                        <Users className="text-slate-300" style={{ width: 'var(--fluid-space-md)', height: 'var(--fluid-space-md)' }} />
                        <p className="font-black text-slate-400 uppercase tracking-widest italic" style={{ fontSize: 'calc(var(--fluid-text-xs) * 0.7)' }}>Vacío</p>
                      </div>
                    )}
                  </div>
                </CardContent>
              </Card>
            </motion.div>
          ))}
        </div>
      </div>
    </motion.div>
  );

  const renderGame = () => (
    <div 
      className="flex w-full h-full mx-auto overflow-hidden min-h-0 items-stretch relative bg-slate-50/20"
      style={{ 
        maxWidth: 'min(1700px, 100.0%)', 
        padding: 'var(--fluid-space-md)',
        gap: 'var(--fluid-space-md)'
      }}
    >
      {/* Botón para Terminar el Juego - Altamente Visible en Esquina Superior Derecha */}
      <div 
        className="absolute z-50 pointer-events-auto" 
        style={{ 
          top: '12px', 
          right: '12px'
        }}
      >
        <Button 
          variant="destructive"
          onClick={endGame}
          className="rounded-xl shadow-[0_8px_24px_rgba(239,68,68,0.4)] bg-gradient-to-r from-red-600 to-rose-600 hover:from-red-700 hover:to-rose-700 text-white font-black uppercase tracking-wider flex items-center justify-center gap-1.5 border-2 border-white active:scale-95 transition-all text-[11px] h-9 px-3.5"
        >
          <LogOut className="animate-pulse" style={{ width: '13px', height: '13px' }} />
          <span>TERMINAR JUEGO</span>
        </Button>
      </div>

      {/* COLUMN 1: Teams Sidebar */}
      <div 
        className="hidden lg:flex flex-col glass-panel rounded-xl shrink-0 overflow-hidden shadow-2xl bg-white/70 min-h-0"
        style={{ width: 'clamp(380px, 26vw, 480px)', padding: 'var(--fluid-space-md)' }}
      >
        <div className="text-center" style={{ marginBottom: 'var(--fluid-space-md)' }}>
          <h3 className="text-slate-800 font-black tracking-[0.2em] opacity-80 uppercase leading-none" style={{ fontSize: 'var(--fluid-text-sm)' }}>Equipos</h3>
          <p className="font-bold text-slate-400 uppercase tracking-[0.4em] mt-1" style={{ fontSize: 'calc(var(--fluid-text-xs) * 0.8)' }}>TabúEdu Pro</p>
        </div>
        <div className="overflow-y-auto custom-scrollbar flex-1 pr-1" style={{ gap: 'var(--fluid-space-xs)', display: 'flex', flexDirection: 'column' }}>
          {[...Array(numTeams).keys()]
            .sort((a, b) => {
              if (a === currentTeamIndex) return -1;
              if (b === currentTeamIndex) return 1;
              return a - b;
            })
            .map((idx) => {
              const score = teamScores[idx];
              const roundsTaken = getTeamRound(idx);
              const isActive = idx === currentTeamIndex;
              
              return (
                <motion.div 
                  key={idx} 
                  layout
                  initial={false}
                  animate={{ 
                    scale: isActive ? 1.02 : 0.98,
                    opacity: isActive ? 1 : 0.6,
                    zIndex: isActive ? 10 : 1
                  }}
                  transition={{ type: "spring", stiffness: 300, damping: 30 }}
                  className={`rounded-lg transition-all duration-500 flex flex-col border-2 ${
                    isActive 
                      ? `${TEAM_COLORS[idx].lightBg} ${TEAM_COLORS[idx].border} shadow-xl ring-2 ring-white` 
                      : 'bg-slate-100/50 border-transparent grayscale-[0.2]'
                  }`}
                  style={{ padding: 'var(--fluid-space-xs)', gap: 'var(--fluid-space-xs)' }}
                >
                  <div className="flex items-center justify-between" style={{ gap: 'var(--fluid-space-xs)' }}>
                    <div className="flex-1 text-center group/score flex flex-col items-center">
                      <div className="flex items-center" style={{ gap: '4px' }}>
                        <button 
                          onClick={() => setTeamScores(prev => {
                            const next = [...prev];
                            next[idx] = Math.max(0, next[idx] - 1);
                            return next;
                          })}
                          className="opacity-0 group-hover/score:opacity-100 transition-opacity text-slate-400 hover:text-rose-500 p-0.5 hover:bg-rose-50 rounded"
                        >
                          <Minus size={10} />
                        </button>
                        <p className={`font-black leading-none tabular-nums ${isActive ? TEAM_COLORS[idx].textColor : 'text-slate-400'}`} style={{ fontSize: 'var(--fluid-text-sm)' }}>{score}</p>
                        <button 
                          onClick={() => setTeamScores(prev => {
                            const next = [...prev];
                            next[idx] = next[idx] + 1;
                            return next;
                          })}
                          className="opacity-0 group-hover/score:opacity-100 transition-opacity text-slate-400 hover:text-emerald-500 p-0.5 hover:bg-emerald-50 rounded"
                        >
                          <Plus size={10} />
                        </button>
                      </div>
                      <p className="font-bold text-slate-400 uppercase tracking-tighter mt-0.5" style={{ fontSize: 'calc(var(--fluid-text-xs) * 0.6)' }}>Puntos</p>
                    </div>
                    
                    <div className="flex-1 text-center">
                      <p className={`font-black leading-none tabular-nums ${isActive ? TEAM_COLORS[idx].textColor : 'text-slate-400'}`} style={{ fontSize: 'var(--fluid-text-sm)' }}>{roundsTaken}/{totalRounds}</p>
                      <p className="font-bold text-slate-400 uppercase tracking-tighter mt-0.5" style={{ fontSize: 'calc(var(--fluid-text-xs) * 0.6)' }}>Rondas</p>
                    </div>
                  </div>
                  <div className="flex flex-col items-center">
                    <div className="flex items-center gap-2 mb-1.5">
                      <div className={`w-3 h-3 rounded-full ${TEAM_COLORS[idx].bg}`} />
                      <p className={`font-black truncate uppercase tracking-widest text-center leading-none ${isActive ? TEAM_COLORS[idx].textColor : 'text-slate-400'}`} style={{ fontSize: 'calc(var(--fluid-text-xs) * 0.7)' }}>
                        {TEAM_COLORS[idx].name}
                      </p>
                      <div className={`w-3 h-3 rounded-full ${TEAM_COLORS[idx].bg}`} />
                    </div>
                    <div className="flex flex-wrap justify-center overflow-hidden w-full" style={{ gap: '4px' }}>
                      {(teamMembers[idx] || []).map((member, mIdx) => (
                        <span 
                          key={mIdx} 
                          className={`font-black px-2.5 py-1 rounded-[4px] uppercase tracking-tighter ${isActive ? 'bg-white text-slate-900 shadow-sm border border-slate-100' : 'bg-slate-200/60 text-slate-500'} inline-block`}
                          style={{ fontSize: 'calc(var(--fluid-text-xs) * 0.95)', lineHeight: '1' }}
                        >
                          {member.name}
                        </span>
                      ))}
                    </div>
                  </div>
                  {isActive && (
                    <motion.div 
                      layoutId="active-indicator"
                      className="h-0.5 w-full bg-current opacity-20 rounded-full"
                      style={{ marginTop: 'calc(var(--fluid-space-xs) * 0.5)' }}
                    />
                  )}
                </motion.div>
              );
            })}
        </div>
        <div className="border-t border-slate-100 opacity-20 text-center" style={{ marginTop: 'var(--fluid-space-xs)', paddingTop: 'var(--fluid-space-xs)' }}>
          <p className="uppercase font-black tracking-widest text-slate-900 leading-tight" style={{ fontSize: 'calc(var(--fluid-text-xs) * 0.6)' }}>
            I.E. José Andrés Rázuri<br/>Pacasmayo
          </p>
        </div>
      </div>

      {/* COLUMN 2: Main Play Area */}
      <div className="flex-1 flex flex-col min-w-0 overflow-hidden h-full items-center" style={{ gap: 'var(--fluid-space-sm)' }}>
        {/* Header */}
        <div 
          className="glass-panel rounded-xl flex items-center justify-between shrink-0 shadow-lg border-slate-100 bg-white/70 w-full"
          style={{ padding: 'var(--fluid-space-xs)', gap: 'var(--fluid-space-sm)', maxWidth: 'min(25rem, 100%)' }}
        >
          <div className="flex items-center" style={{ gap: 'var(--fluid-space-xs)' }}>
            <div 
              className={`rounded-full flex items-center justify-center border-2 shadow-md ${TEAM_COLORS[currentTeamIndex].bg} ${TEAM_COLORS[currentTeamIndex].border}`}
              style={{ width: 'var(--fluid-space-md)', height: 'var(--fluid-space-md)' }}
            />
            <div className="flex flex-col">
              <p className="font-black text-slate-400 uppercase tracking-widest leading-none mb-1" style={{ fontSize: 'calc(var(--fluid-text-xs) * 0.7)' }}>Turno {TEAM_COLORS[currentTeamIndex].name}</p>
              <div className="flex flex-col gap-0.5">
                <span className="text-[11px] font-black uppercase text-emerald-600 leading-tight">
                  📢 Pistas: <span className="text-slate-800">{activeCommunicators.map(c => c.name).join(' y ') || 'Cualquiera'}</span>
                </span>
                <span className="text-[11px] font-black uppercase text-orange-600 leading-tight">
                  🎯 Adivina: <span className="text-slate-800">{getGuesserNames() || 'Resto del equipo'}</span>
                </span>
              </div>
            </div>
          </div>
          
          <div className="flex-1 h-1 bg-slate-100 rounded-full overflow-hidden hidden md:block">
            <motion.div 
              className="h-full bg-blue-500 opacity-60"
              initial={{ width: 0 }}
              animate={{ width: `${(turnCounter / (numTeams * totalRounds)) * 100}%` }}
              transition={{ duration: 0.5 }}
            />
          </div>

          <div className="text-right">
            <p className="font-black text-slate-400 uppercase tracking-widest leading-none mb-1" style={{ fontSize: 'calc(var(--fluid-text-xs) * 0.7)' }}>Puntaje</p>
            <h4 className="text-slate-900 font-black leading-none tabular-nums" style={{ fontSize: 'var(--fluid-text-xl)' }}>{teamScores[currentTeamIndex]}</h4>
          </div>
        </div>

        {/* Card Section */}
        <div className="flex-1 min-h-0 relative w-full mx-auto" style={{ maxWidth: 'min(25rem, 100%)' }}>
          <AnimatePresence mode="wait">
            {currentCard && (
              <motion.div
                key={currentCard.secreto}
                initial={{ opacity: 0, scale: 0.98, y: 5 }}
                animate={{ opacity: 1, scale: 1, y: 0 }}
                exit={{ opacity: 0, scale: 1.02, y: -5 }}
                transition={{ type: "spring", damping: 30, stiffness: 250 }}
                className="absolute inset-0 flex flex-col"
              >
                <div className="flex-1 bg-white rounded-[var(--radius-3xl)] overflow-hidden shadow-2xl flex flex-col border-[4px] border-white/5 w-full mx-auto">
                  {/* Secreto block */}
                  <div 
                    className={`${TEAM_COLORS[currentTeamIndex].bg} text-center border-b-4 ${TEAM_COLORS[currentTeamIndex].border} flex flex-col items-center justify-center relative overflow-hidden group shrink-0`}
                    style={{ minHeight: 'clamp(100px, 14vh, 130px)', padding: 'var(--fluid-space-sm)' }}
                  >
                    <div className="absolute inset-0 bg-black/5 opacity-0 group-hover:opacity-100 transition-opacity" />
                    <div className="font-black text-white/95 uppercase tracking-[0.12em] mb-1 z-10 flex items-center justify-center gap-1 my-1" style={{ fontSize: 'calc(var(--fluid-text-xs) * 0.72)' }}>
                      <span>PALABRA SECRETA • NIVEL</span>
                      <span className={`px-2 py-0.5 rounded-full text-[9px] font-black shadow-lg ${
                        currentCard.nivelLogro === "Destacado" ? "bg-amber-400 text-slate-950 border border-amber-300" :
                        currentCard.nivelLogro === "Logrado" ? "bg-emerald-400 text-slate-900 border border-emerald-300" :
                        currentCard.nivelLogro === "Proceso" ? "bg-sky-400 text-slate-900 border border-sky-300" :
                        "bg-slate-100 text-slate-800 border border-slate-200"
                      }`}>
                        {currentCard.nivelLogro || "Inicio"}
                      </span>
                    </div>
                    <h2 
                      className="text-white font-black tracking-tighter uppercase leading-[0.9] drop-shadow-xl z-10 text-center break-words w-full flex flex-col items-center justify-center px-4"
                      style={{ 
                        fontSize: currentCard.secreto.length <= 8 
                          ? 'var(--fluid-text-4xl)' 
                          : currentCard.secreto.length <= 15 
                          ? 'var(--fluid-text-3xl)' 
                          : 'var(--fluid-text-2xl)'
                      }}
                    >
                      {(() => {
                        const words = currentCard.secreto.split(' ');
                        // Don't split if short (2 words and total length <= 16)
                        if (words.length <= 2 && currentCard.secreto.length <= 16) {
                          return <span>{currentCard.secreto}</span>;
                        }
                        if (words.length <= 2) {
                          return words.map((part, i) => (
                            <span key={i} className="block leading-none">{part}</span>
                          ));
                        }
                        const mid = Math.ceil(words.length / 2);
                        return [
                          words.slice(0, mid).join(' '),
                          words.slice(mid).join(' ')
                        ].map((part, i) => (
                          <span key={i} className="block leading-none">{part}</span>
                        ));
                      })()}
                    </h2>
                  </div>
                  {/* Taboo block */}
                  <div 
                    className="flex-1 flex flex-col items-center justify-center overflow-hidden" 
                    style={{ padding: 'var(--fluid-space-md) var(--fluid-space-sm)', gap: 'var(--fluid-space-lg)' }}
                  >
                    {currentCard.tabu.map((word, i) => (
                      <React.Fragment key={i}>
                        <p 
                          className="text-slate-900 font-black uppercase text-center tracking-normal leading-[1.1] break-words w-full flex flex-col items-center px-4"
                          style={{ 
                            fontSize: word.length <= 10 
                              ? 'calc(var(--fluid-text-2xl) + 7px)' 
                              : word.length <= 15 
                              ? 'calc(var(--fluid-text-xl) + 7px)' 
                              : 'calc(var(--fluid-text-lg) + 7px)' 
                          }}
                        >
                          {(() => {
                            const words = word.split(' ');
                            if (words.length <= 2 && word.length <= 16) {
                              return <span>{word}</span>;
                            }
                            if (words.length <= 2) {
                              return words.map((part, j) => (
                                <span key={j} className="block leading-[1.05]">{part}</span>
                              ));
                            }
                            const mid = Math.ceil(words.length / 2);
                            return [
                              words.slice(0, mid).join(' '),
                              words.slice(mid).join(' ')
                            ].map((part, j) => (
                              <span key={j} className="block leading-tight">{part}</span>
                            ));
                          })()}
                        </p>
                        {i < currentCard.tabu.length - 1 && (
                          <div className="w-1/4 h-[2px] bg-slate-100 rounded-full shrink-0" />
                        )}
                      </React.Fragment>
                    ))}
                  </div>
                </div>
              </motion.div>
            )}
          </AnimatePresence>
        </div>

        {/* Control Buttons row styled as tactile youth-appealing console action buttons */}
        <div 
          className="grid grid-cols-3 shrink-0 mx-auto w-full" 
          style={{ gap: 'var(--fluid-space-xs)', paddingBottom: 'var(--fluid-space-xs)', maxWidth: 'min(25rem, 100%)' }}
        >
          <motion.div whileHover={{ scale: 1.05 }} whileTap={{ scale: 0.94 }} className="w-full">
            <Button 
              className="w-full rounded-2xl shadow-[0_4px_12px_rgba(239,68,68,0.25)] active:shadow-none transition-all bg-gradient-to-br from-rose-500 via-red-500 to-rose-600 hover:from-rose-600 hover:to-rose-700 border-none group"
              style={{ height: 'calc(var(--fluid-space-xl) + 6px)' }}
              onClick={handleTaboo}
              disabled={!currentCard || !isTimerActive}
            >
              <X className="text-white stroke-[4] group-hover:rotate-90 transition-transform duration-300" style={{ width: 'var(--fluid-space-md)', height: 'var(--fluid-space-md)' }} />
            </Button>
          </motion.div>
          
          <motion.div whileHover={{ scale: 1.05 }} whileTap={{ scale: 0.94 }} className="w-full">
            <Button 
              className="w-full rounded-2xl shadow-[0_4px_12px_rgba(79,70,229,0.25)] active:shadow-none transition-all bg-gradient-to-br from-indigo-500 via-blue-500 to-indigo-600 hover:from-indigo-600 hover:to-indigo-700 border-none group"
              style={{ height: 'calc(var(--fluid-space-xl) + 6px)' }}
              onClick={handlePass}
              disabled={!currentCard || !isTimerActive || passesUsedInTurn >= passesPerTurn}
            >
              <SkipForward className="text-white fill-current group-hover:translate-x-1 transition-transform" style={{ width: 'var(--fluid-space-md)', height: 'var(--fluid-space-md)' }} />
            </Button>
          </motion.div>

          <motion.div whileHover={{ scale: 1.05 }} whileTap={{ scale: 0.94 }} className="w-full">
            <Button 
              className="w-full rounded-2xl shadow-[0_4px_12px_rgba(16,185,129,0.25)] active:shadow-none transition-all bg-gradient-to-br from-emerald-400 via-teal-500 to-emerald-600 hover:from-emerald-500 hover:to-emerald-600 border-none group"
              style={{ height: 'calc(var(--fluid-space-xl) + 6px)' }}
              onClick={handleCorrect}
              disabled={!currentCard || !isTimerActive}
            >
              <Check className="text-white stroke-[4] group-hover:scale-125 transition-transform" style={{ width: 'var(--fluid-space-md)', height: 'var(--fluid-space-md)' }} />
            </Button>
          </motion.div>
        </div>
      </div>

      {/* COLUMN 3: Timer & Meta Sidebar */}
      <div 
        className="hidden lg:flex flex-col shrink-0 h-full overflow-hidden min-h-0" 
        style={{ width: 'clamp(380px, 26vw, 480px)', gap: 'var(--fluid-space-md)' }}
      >
        {/* Institutional Top */}
        <div className="text-center opacity-30" style={{ padding: '0 var(--fluid-space-md)', marginTop: 'var(--fluid-space-xs)' }}>
          <h5 className="font-black text-slate-800 uppercase tracking-[0.2em] leading-tight" style={{ fontSize: 'calc(var(--fluid-text-xs) * 0.8)' }}>
            UGEL Pacasmayo<br/>San Pedro de Lloc
          </h5>
        </div>

        {/* Timer Box */}
        <div 
          className="glass-panel rounded-xl flex flex-col items-center justify-center shrink-0 shadow-lg bg-white/70"
          style={{ padding: 'var(--fluid-space-md)' }}
        >
          <div className="relative flex items-center justify-center" style={{ width: 'clamp(120px, 15vh, 160px)', height: 'clamp(120px, 15vh, 160px)' }}>
            <svg className="absolute inset-0 w-full h-full -rotate-90">
              <circle 
                cx="50%" cy="50%" r="45%" 
                fill="transparent" 
                stroke="rgba(0,0,0,0.05)" 
                strokeWidth="10" 
              />
              <motion.circle 
                cx="50%" cy="50%" r="45%" 
                fill="transparent" 
                stroke="currentColor" 
                strokeWidth="6" 
                className={`${timeLeft <= 10 ? 'text-rose-500' : 'text-blue-500'} transition-colors duration-500`}
                initial={{ strokeDasharray: "283 283" }}
                animate={{ strokeDasharray: `${(timeLeft / timePerTurn) * 283} 283` }}
                strokeLinecap="round"
              />
            </svg>
            <div className="text-center z-10 flex flex-col items-center">
              <span className={`font-bold uppercase tracking-widest text-slate-400`} style={{ fontSize: 'calc(var(--fluid-text-xs) * 0.7)' }}>Restante</span>
              <span className={`font-black tabular-nums transition-colors ${timeLeft <= 10 ? 'text-rose-600' : 'text-slate-900'}`} style={{ fontSize: 'var(--fluid-text-2xl)' }}>{timeLeft}</span>
            </div>
          </div>
        </div>

        {/* AI Hint Section */}
        <div className="glass-panel rounded-xl flex-1 flex flex-col bg-white/70 shadow-xl overflow-hidden" style={{ padding: 'var(--fluid-space-sm)' }}>
          <div className="flex items-center justify-between" style={{ marginBottom: 'var(--fluid-space-xs)' }}>
            <div className="flex items-center" style={{ gap: 'var(--fluid-space-xs)' }}>
              <Sparkles className="text-amber-500" style={{ width: 'var(--fluid-space-xs)', height: 'var(--fluid-space-xs)' }} />
              <h4 className="font-black text-slate-800 uppercase tracking-widest" style={{ fontSize: 'calc(var(--fluid-text-xs) * 0.8)' }}>Socrática IA</h4>
            </div>
            <Button 
              variant="ghost" 
              size="icon" 
              className="h-6 w-6 text-slate-300 hover:text-slate-500 rounded-full transition-colors"
              onClick={() => setIsHintsVisible(!isHintsVisible)}
              title={isHintsVisible ? "Ocultar pistas" : "Mostrar pistas"}
            >
              {isHintsVisible ? <Eye size={14} /> : <EyeOff size={14} />}
            </Button>
          </div>
          
          <div className="flex-1 bg-slate-50/50 rounded-lg p-3 custom-scrollbar overflow-y-auto border border-slate-100 flex flex-col">
            {!isHintsVisible ? (
              <div className="flex-1 flex flex-col items-center justify-center text-center opacity-40">
                <EyeOff className="text-slate-300 mb-2" style={{ width: 'var(--fluid-space-md)', height: 'var(--fluid-space-md)' }} />
                <p className="font-black text-slate-400 uppercase tracking-widest italic" style={{ fontSize: 'calc(var(--fluid-text-xs) * 0.7)' }}>Pistas Bloqueadas</p>
              </div>
            ) : (
              <>
                {!socraticHint && !isHintLoading && (
                  <div className="flex-1 flex flex-col items-center justify-center text-center">
                    <Info className="text-slate-300 mb-2" style={{ width: 'var(--fluid-space-md)', height: 'var(--fluid-space-md)' }} />
                    <p className="font-black text-slate-400 uppercase tracking-tighter italic" style={{ fontSize: 'calc(var(--fluid-text-xs) * 0.7)' }}>Buscando pista...</p>
                  </div>
                )}
                
                {isHintLoading && (
                  <div className="flex-1 flex flex-col items-center justify-center space-y-2">
                    <RotateCcw className="animate-spin text-blue-500" style={{ width: 'var(--fluid-space-sm)', height: 'var(--fluid-space-sm)' }} />
                    <p className="font-black text-slate-400 uppercase italic animate-pulse" style={{ fontSize: 'calc(var(--fluid-text-xs) * 0.7)' }}>Interpretando...</p>
                  </div>
                )}

                {socraticHint && (
                  <motion.div 
                    initial={{ opacity: 0, y: 10 }} 
                    animate={{ opacity: 1, y: 0 }}
                    key={currentCard?.secreto}
                    className="flex-1 flex flex-col justify-center"
                  >
                    <p className="font-bold text-blue-600 uppercase tracking-tighter mb-2" style={{ fontSize: 'calc(var(--fluid-text-xs) * 0.9)' }}>Pista para los participantes:</p>
                    <p className="font-extrabold text-slate-800 leading-snug italic text-center text-blue-900" style={{ fontSize: 'var(--fluid-text-2xl)' }}>"{socraticHint}"</p>
                    <div className="mt-4 pt-4 border-t border-slate-100">
                       <p className="font-black text-[7px] text-slate-300 uppercase tracking-widest text-center">Pista generada automáticamente</p>
                    </div>
                  </motion.div>
                )}
              </>
            )}
          </div>
        </div>

        {/* Footer info */}
        <div className="p-2 bg-slate-900 rounded-xl shadow-lg shrink-0">
           <div className="flex items-center justify-center" style={{ gap: 'var(--fluid-space-xs)' }}>
              <School className="text-blue-400" style={{ width: 'var(--fluid-space-xs)', height: 'var(--fluid-space-xs)' }} />
              <span className="font-black text-white uppercase tracking-tighter" style={{ fontSize: 'calc(var(--fluid-text-xs) * 0.6)' }}>Rázuri Educativo</span>
           </div>
        </div>
      </div>

      {/* Global Back Button during Game (for all states except when modal blocks it) */}
      <motion.div 
        className="absolute top-3 left-3 z-[1000] pointer-events-auto"
        animate={{ opacity: [0.4, 1, 0.4] }}
        transition={{ repeat: Infinity, duration: 2, ease: "easeInOut" }}
      >
        <Button 
          variant="ghost"
          className="h-9 w-9 p-0 rounded-full text-slate-400 hover:text-slate-600 hover:bg-white/50 backdrop-blur-sm transition-all shadow-sm border border-transparent hover:border-slate-200"
          title="Regresar a configuración"
          onClick={() => {
            if (window.confirm("¿Estás seguro de que quieres salir del juego actual? Se perderá el progreso del turno.")) {
              setShowTurnModal(false);
              setScreen("start");
            }
          }}
        >
          <ChevronLeft size={20} />
        </Button>
      </motion.div>

      {/* Turn Modal */}
      <Dialog open={showTurnModal} onOpenChange={(open) => { if (open) setShowTurnModal(true); }}>
        <DialogContent 
          showCloseButton={false}
          className="text-center rounded-[var(--radius-3xl)] border-none glass-panel shadow-[0_30px_60px_-10px_rgba(0,0,0,0.1)] text-slate-900 bg-white/90 overflow-hidden"
          style={{ width: 'min(30rem, 90vw)', padding: 'var(--fluid-space-md)' }}
        >
          <DialogHeader className="space-y-1 relative">
            <motion.div
              className="absolute -top-2 -left-2 z-50"
              animate={{ opacity: [0.3, 0.8, 0.3] }}
              transition={{ repeat: Infinity, duration: 2, ease: "easeInOut" }}
            >
              <Button
                variant="ghost"
                className="h-8 w-8 p-0 rounded-full text-slate-300 hover:text-slate-500 hover:bg-slate-100 hover:opacity-100 transition-all"
                onClick={() => {
                  setShowTurnModal(false);
                  setScreen("start");
                }}
                title="Regresar"
              >
                <ChevronLeft size={16} />
              </Button>
            </motion.div>
            <DialogTitle className="font-black mb-1 uppercase italic tracking-tighter text-slate-900" style={{ fontSize: 'var(--fluid-text-xl)' }}>¡Prepárate!</DialogTitle>
            <DialogDescription className="font-bold text-slate-500 uppercase tracking-widest leading-tight" style={{ fontSize: 'var(--fluid-text-xs)' }}>
              Es el turno del <br/>
              <span 
                className={`inline-block mt-1 font-black px-4 py-1 rounded-xl shadow-xl uppercase tracking-tighter ${TEAM_COLORS[currentTeamIndex].bg} text-white border-b-4 ${TEAM_COLORS[currentTeamIndex].border}`}
                style={{ fontSize: 'var(--fluid-text-lg)' }}
              >
                {TEAM_COLORS[currentTeamIndex].name}
              </span>
              <div className="mt-2.5 p-3 bg-white/70 backdrop-blur-md rounded-xl border border-slate-200 text-left flex flex-col gap-2">
                <div>
                  <p className="font-black text-emerald-600 uppercase tracking-widest mb-0.5" style={{ fontSize: '10px' }}>📢 COMUNICA PISTAS (Sorteado al azar):</p>
                  <p className="font-extrabold text-slate-800 uppercase tracking-tight text-sm">
                    {activeCommunicators.map(c => c.name).join(' y ') || 'Cualquiera'}
                  </p>
                </div>
                <div className="border-t border-slate-100 pt-1.5">
                  <p className="font-black text-orange-600 uppercase tracking-widest mb-0.5" style={{ fontSize: '10px' }}>🎯 DEBE ADIVINAR:</p>
                  <p className="font-extrabold text-slate-800 uppercase tracking-tight text-sm">
                    {getGuesserNames() || 'Resto del equipo'}
                  </p>
                </div>
              </div>
            </DialogDescription>
          </DialogHeader>
          <div className="py-4">
            <div 
              className={`rounded-full mx-auto flex items-center justify-center font-black text-white shadow-xl border-[4px] ${TEAM_COLORS[currentTeamIndex].bg} ${TEAM_COLORS[currentTeamIndex].border}`}
              style={{ width: 'var(--fluid-space-xl-3)', height: 'var(--fluid-space-xl-3)', fontSize: 'var(--fluid-text-hero)' }}
            >
              {currentTeamIndex + 1}
            </div>
          </div>
          <DialogFooter className="sm:justify-center">
            <Button 
              className="w-full font-black rounded-xl shadow-xl bg-blue-600 text-white hover:bg-blue-700 uppercase tracking-[0.2em]" 
              style={{ height: 'var(--fluid-space-xl)', fontSize: 'var(--fluid-text-base)' }}
              onClick={startTurn}
            >
              COMENZAR
            </Button>
          </DialogFooter>
        </DialogContent>
      </Dialog>
    </div>
  );

  useEffect(() => {
    if (screen === 'stats') {
      const duration = 5 * 1000;
      const animationEnd = Date.now() + duration;
      const defaults = { startVelocity: 30, spread: 360, ticks: 60, zIndex: 0 };

      const randomInRange = (min: number, max: number) => Math.random() * (max - min) + min;

      const interval: any = setInterval(function() {
        const timeLeft = animationEnd - Date.now();

        if (timeLeft <= 0) {
          return clearInterval(interval);
        }

        const particleCount = 50 * (timeLeft / duration);
        confetti({ ...defaults, particleCount, origin: { x: randomInRange(0.1, 0.3), y: Math.random() - 0.2 } });
        confetti({ ...defaults, particleCount, origin: { x: randomInRange(0.7, 0.9), y: Math.random() - 0.2 } });
      }, 250);

      const end = Date.now() + 500;

      (function frame() {
        confetti({
          particleCount: 2,
          angle: 60,
          spread: 55,
          origin: { x: 0 },
          colors: ['#3b82f6', '#f59e0b', '#ef4444']
        });
        confetti({
          particleCount: 2,
          angle: 120,
          spread: 55,
          origin: { x: 1 },
          colors: ['#3b82f6', '#f59e0b', '#ef4444']
        });

        if (Date.now() < end) {
          requestAnimationFrame(frame);
        }
      }());
      
      return () => clearInterval(interval);
    }
  }, [screen]);

  const renderStats = () => {
    const sortedTeams = teamScores
      .map((score, index) => ({ score, index }))
      .sort((a, b) => b.score - a.score);
    
    return (
      <motion.div 
        initial={{ opacity: 0, scale: 0.98 }}
        animate={{ opacity: 1, scale: 1 }}
        className="w-full mx-auto text-center flex flex-col justify-center min-h-0 h-full py-2 relative"
        style={{ maxWidth: 'min(80rem, 100%)', padding: 'var(--fluid-space-xs)' }}
      >
        <div className="absolute top-4 right-4 z-[100] flex gap-2">
          <Button 
            variant="outline"
            className="h-10 w-10 p-0 rounded-full bg-white/20 hover:bg-white/40 border-white/30 text-slate-800 shadow-sm"
            title="Nueva Configuración"
            onClick={() => setScreen("start")}
          >
            <Settings size={18} />
          </Button>
          <Button 
            variant="outline"
            className="h-10 w-10 p-0 rounded-full bg-blue-500 hover:bg-blue-600 border-none text-white shadow-md animate-pulse"
            title="Reintentar Juego"
            onClick={() => {
              setTeamScores(Array(numTeams).fill(0));
              setTurnCounter(0);
              setUsedWords([]);
              setTeamDiffCycles(Array(numTeams).fill(0)); // Reset cycles
              setTimeLeft(timePerTurn);
              setScreen("game");
              setShowTurnModal(true);
            }}
          >
            <RotateCcw size={18} />
          </Button>
          <Button 
            variant="outline"
            className="h-10 px-4 rounded-full bg-emerald-500 hover:bg-emerald-600 border-none text-white shadow-md flex items-center gap-2"
            title="Exportar a Word"
            onClick={async () => {
              const doc = new Document({
                sections: [{
                  properties: {},
                  children: [
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: "Reporte de Resultados - Juego de Tabú Profesional",
                          bold: true,
                          size: 32,
                        }),
                      ],
                      alignment: AlignmentType.CENTER,
                      spacing: { after: 400 },
                    }),
                    new Table({
                      width: {
                        size: 100,
                        type: WidthType.PERCENTAGE,
                      },
                      rows: [
                        new TableRow({
                          children: [
                            new TableCell({
                              children: [new Paragraph({ children: [new TextRun({ text: "Posición", bold: true })], alignment: AlignmentType.CENTER })],
                              verticalAlign: VerticalAlign.CENTER,
                            }),
                            new TableCell({
                              children: [new Paragraph({ children: [new TextRun({ text: "Equipo", bold: true })], alignment: AlignmentType.CENTER })],
                              verticalAlign: VerticalAlign.CENTER,
                            }),
                            new TableCell({
                              children: [new Paragraph({ children: [new TextRun({ text: "Puntaje", bold: true })], alignment: AlignmentType.CENTER })],
                              verticalAlign: VerticalAlign.CENTER,
                            }),
                            new TableCell({
                              children: [new Paragraph({ children: [new TextRun({ text: "Integrantes", bold: true })], alignment: AlignmentType.CENTER })],
                              verticalAlign: VerticalAlign.CENTER,
                            }),
                          ],
                        }),
                        ...sortedTeams.map((team, idx) => (
                          new TableRow({
                            children: [
                              new TableCell({
                                children: [new Paragraph({ text: `${idx + 1}º`, alignment: AlignmentType.CENTER })],
                              }),
                              new TableCell({
                                children: [new Paragraph({ text: TEAM_COLORS[team.index].name, alignment: AlignmentType.CENTER })],
                              }),
                              new TableCell({
                                children: [new Paragraph({ text: team.score.toString(), alignment: AlignmentType.CENTER })],
                              }),
                              new TableCell({
                                children: [new Paragraph({ text: (teamMembers[team.index] || []).map(m => m.name).join(', ') })],
                              }),
                            ],
                          })
                        )),
                      ],
                    }),
                    new Paragraph({
                      children: [
                        new TextRun({
                          text: `Documento generado el ${new Date().toLocaleDateString()} ${new Date().toLocaleTimeString()}`,
                          italics: true,
                          size: 18,
                        }),
                      ],
                      alignment: AlignmentType.RIGHT,
                      spacing: { before: 400 },
                    }),
                  ],
                }],
              });

              const blob = await Packer.toBlob(doc);
              saveAs(blob, `Resultados_Tabu_${new Date().toLocaleDateString().replace(/\//g, '-')}.docx`);
              toast.success("Documento Word generado con éxito");
            }}
          >
            <FileText size={18} />
            <span className="hidden sm:inline font-bold">EXPORTAR PARA REGISTRO</span>
          </Button>
        </div>

        <div className="relative shrink-0 mb-2 pt-6">
          <motion.div 
            initial={{ scale: 0, rotate: -45 }}
            animate={{ scale: 1, rotate: 0 }}
            transition={{ type: "spring", damping: 12, delay: 0.5 }}
            className="bg-gradient-to-tr from-amber-400 to-yellow-300 rounded-full w-fit mx-auto border-4 border-white shadow-[0_0_40px_rgba(251,191,36,0.6)] relative mb-3"
            style={{ padding: '12px' }}
          >
            <Crown className="text-white w-10 h-10 relative z-10 drop-shadow-md" />
            <motion.div 
              animate={{ scale: [1, 1.8, 1], opacity: [0.5, 0, 0.5] }}
              transition={{ repeat: Infinity, duration: 2.5 }}
              className="absolute inset-0 bg-amber-400 rounded-full"
            />
          </motion.div>
          <h2 className="font-black tracking-tighter text-slate-900 italic uppercase leading-none drop-shadow-xl mb-2" style={{ fontSize: 'var(--fluid-text-xl-4)' }}>¡Victoria Magistral!</h2>
          <p className="text-slate-500 font-extrabold uppercase tracking-[0.6em]" style={{ fontSize: '10px' }}>Ceremonia de Premiación</p>
        </div>

        {/* The Real Podium */}
        <div 
          className="flex flex-row justify-center items-end w-full max-w-7xl mx-auto px-4 mb-6 overflow-visible relative" 
          style={{ gap: '16px' }}
        >
          {/* Creative Background Aura for the Winner Area */}
          <div className="absolute inset-0 pointer-events-none -z-10 flex justify-center items-center">
             <motion.div 
               animate={{ 
                 scale: [1, 1.2, 1],
                 rotate: [0, 90, 180, 270, 360]
               }}
               transition={{ duration: 15, repeat: Infinity, ease: "linear" }}
               className="w-[60vw] h-[60vw] max-w-[800px] max-h-[800px] opacity-10 bg-[conic-gradient(from_0deg,transparent,rgba(59,130,246,0.3),transparent,rgba(251,191,36,0.3),transparent)] blur-3xl"
             />
          </div>

          {(() => {
            const n = sortedTeams.length;
            const centerIdx = Math.floor((n - 1) / 2);
            const reordered = new Array(n);
            let left = centerIdx - 1;
            let right = centerIdx + 1;
            reordered[centerIdx] = sortedTeams[0];
            for (let i = 1; i < n; i++) {
              if (i % 2 === 1) {
                if (left >= 0) reordered[left--] = sortedTeams[i];
                else reordered[right++] = sortedTeams[i];
              } else {
                if (right < n) reordered[right++] = sortedTeams[i];
                else reordered[left--] = sortedTeams[i];
              }
            }

            const uniqueScores = Array.from(new Set(sortedTeams.map(t => t.score)));
            
            return reordered.filter(Boolean).map((team, pIdx) => {
              const rank = sortedTeams.findIndex(t => t.index === team.index);
              // Calculate visual tier based on score to handle ties in height
              const visualTier = uniqueScores.indexOf(team.score);
              const isWinner = visualTier === 0;
              const delay = 0.2 + (rank * 0.15);
              
              const hClass = 
                visualTier === 0 ? "h-52 sm:h-64" : 
                visualTier === 1 ? "h-36 sm:h-48" : 
                visualTier === 2 ? "h-30 sm:h-40" : 
                visualTier === 3 ? "h-26 sm:h-34" :
                visualTier === 4 ? "h-22 sm:h-28" : "h-18 sm:h-22";

              return (
                <motion.div
                  key={team.index}
                  initial={{ opacity: 0, scale: 0.5, y: 100 }}
                  animate={{ opacity: 1, scale: isWinner ? 1.1 : 1, y: 0 }}
                  transition={{ delay, type: "spring", stiffness: 70, damping: 15 }}
                  className={`flex flex-col items-center flex-1 min-w-0 max-w-[220px] relative ${isWinner ? 'z-20' : 'z-10'}`}
                >
                  {/* Winner Sparkles (Only for top tier) */}
                  {isWinner && (
                    <motion.div 
                      className="absolute -top-16 inset-x-0 flex justify-around pointer-events-none"
                      animate={{ y: [0, -10, 0] }}
                      transition={{ duration: 2, repeat: Infinity }}
                    >
                      <Sparkles className="text-yellow-400 w-8 h-8 drop-shadow-[0_0_10px_rgba(251,191,36,0.8)]" />
                      <Sparkles className="text-yellow-400 w-6 h-6 delay-300 drop-shadow-[0_0_10px_rgba(251,191,36,0.8)]" />
                    </motion.div>
                  )}

                  <motion.div 
                    animate={isWinner ? { y: [0, -8, 0] } : {}}
                    transition={{ duration: 3, repeat: Infinity, ease: "easeInOut" }}
                    className={`rounded-full border-4 mb-3 flex items-center justify-center font-black text-white shadow-[0_10px_30px_rgba(0,0,0,0.2)] ${TEAM_COLORS[team.index].bg} ${TEAM_COLORS[team.index].border} relative`}
                    style={{ width: '48px', height: '48px' }}
                  >
                    {/* Pulsing glow for winner, single flash for others */}
                    <motion.div 
                      className="absolute inset-0 bg-white/40 rounded-full"
                      initial={{ opacity: 0, scale: 0.8 }}
                      animate={isWinner ? { 
                        opacity: [0, 0.4, 0],
                        scale: [0.8, 1.4, 0.8]
                      } : {
                        opacity: [0, 1, 0],
                        scale: [0.8, 1.8, 1.2]
                      }}
                      transition={isWinner ? { 
                        repeat: Infinity, 
                        duration: 3 
                      } : {
                        duration: 1,
                        times: [0, 0.5, 1],
                        delay: delay + 0.5,
                        ease: "easeOut"
                      }}
                    />
                    {visualTier === 0 ? (
                      <Crown size={32} className="drop-shadow-[0_2px_10px_rgba(251,191,36,0.8)] relative z-10" />
                    ) : visualTier === 1 ? (
                      <Medal size={28} className="text-slate-200 drop-shadow-[0_2px_8px_rgba(255,255,255,0.5)] relative z-10" />
                    ) : visualTier === 2 ? (
                      <Award size={26} className="text-orange-300 drop-shadow-[0_2px_8px_rgba(253,186,116,0.5)] relative z-10" />
                    ) : (
                      <Users size={22} className="relative z-10" />
                    )}
                    {isWinner && (
                      <motion.div 
                        animate={{ opacity: [0.3, 0.6, 0.3] }}
                        transition={{ duration: 1.5, repeat: Infinity }}
                        className="absolute -inset-1 bg-white rounded-full -z-10 blur-sm"
                      />
                    )}
                  </motion.div>

                  <div 
                    className={`w-full rounded-t-[2rem] flex flex-col items-center justify-start shadow-2xl relative overflow-hidden border-t-4 border-x border-white/30 backdrop-blur-sm ${TEAM_COLORS[team.index].bg} ${hClass} group transition-all duration-300`}
                  >
                    {/* Glossy & Vibrant Overlays */}
                    <div className="absolute inset-0 bg-gradient-to-tr from-white/20 via-transparent to-black/30 pointer-events-none" />
                    <motion.div 
                      animate={isWinner ? { opacity: [0.1, 0.3, 0.1] } : { opacity: 0.1 }}
                      transition={{ duration: 2, repeat: Infinity }}
                      className="absolute inset-0 bg-white"
                    />
                    
                    <div className="pt-6 px-2 w-full flex flex-col items-center text-white relative z-10 text-center h-full">
                      {/* Removed the redundant color name label to save space and focus on members */}
                      
                      <div className="mt-auto pb-6 flex flex-col items-center w-full px-1">
                        <motion.div 
                          initial={{ scale: 0.8 }}
                          animate={{ scale: 1 }}
                          transition={{ delay: delay + 0.3, type: "spring" }}
                          className="font-black tabular-nums leading-none tracking-tighter drop-shadow-[0_4px_12px_rgba(0,0,0,0.5)] mb-1" 
                          style={{ fontSize: isWinner ? '48px' : (visualTier < 3) ? '36px' : '28px' }}
                        >
                          {team.score}
                        </motion.div>
                        <span className="font-black opacity-90 uppercase tracking-[0.2em] mb-4 drop-shadow-sm" style={{ fontSize: '11px' }}>Puntos</span>
                        
                        {/* Only show students for top 3 tiers */}
                        {visualTier < 3 && (
                          <div className="flex flex-col gap-1 w-full bg-white/15 backdrop-blur-md rounded-xl py-2 px-1 border border-white/20 shadow-inner">
                            <p className="text-[9px] font-black opacity-60 uppercase mb-1 tracking-widest text-white/90">Integrantes</p>
                            <div className="flex flex-col gap-0">
                              {(teamMembers[team.index] || []).slice(0, 5).map((member, i) => (
                                <p key={i} className="text-[13px] font-black uppercase whitespace-nowrap overflow-hidden text-ellipsis leading-tight py-0.5 drop-shadow-md text-white" title={member.name}>
                                  {member.name}
                                </p>
                              ))}
                              {(teamMembers[team.index]?.length || 0) > 5 && (
                                <p className="text-[8px] font-black opacity-50 italic">+{(teamMembers[team.index]?.length || 0) - 5} más</p>
                              )}
                            </div>
                          </div>
                        )}
                      </div>
                    </div>
                  </div>
                </motion.div>
              );
            });
          })()}
        </div>

        {/* Action buttons migrated to top-right to save vertical space */}
      </motion.div>
    );
  };

  return (
    <div className="min-h-screen bg-gradient-to-b from-slate-50 via-indigo-50/20 to-slate-100 text-slate-900 selection:bg-blue-500/10 overflow-hidden flex items-center justify-center p-[var(--fluid-space-xs)] relative" style={{ isolation: 'isolate' }}>
      {/* Beautiful Animated Iridescent Blobs in the Background */}
      <div className="fixed inset-0 overflow-hidden pointer-events-none -z-10">
        <motion.div 
          animate={{
            x: [0, 80, -40, 0],
            y: [0, -60, 50, 0],
            scale: [1, 1.15, 0.9, 1],
          }}
          transition={{
            duration: 25,
            repeat: Infinity,
            ease: "easeInOut"
          }}
          className="absolute -top-[10%] -left-[10%] w-[50dvw] h-[50dvh] rounded-full bg-blue-300/25 blur-[120px]"
        />
        <motion.div 
          animate={{
            x: [0, -100, 60, 0],
            y: [0, 80, -70, 0],
            scale: [1, 0.85, 1.15, 1],
          }}
          transition={{
            duration: 28,
            repeat: Infinity,
            ease: "easeInOut"
          }}
          className="absolute top-[35%] -right-[15%] w-[60dvw] h-[60dvh] rounded-full bg-fuchsia-300/20 blur-[140px]"
        />
        <motion.div 
          animate={{
            x: [0, 50, -60, 0],
            y: [0, 95, -40, 0],
            scale: [1, 1.1, 0.85, 1],
          }}
          transition={{
            duration: 22,
            repeat: Infinity,
            ease: "easeInOut"
          }}
          className="absolute -bottom-[15%] left-[15%] w-[45dvw] h-[45dvh] rounded-full bg-amber-200/20 blur-[110px]"
        />
      </div>

      <Toaster position="top-center" richColors theme="light" />
      
      {/* Main Container limiting dimensions and centering */}
      <main 
        className="relative w-full flex flex-col items-center justify-center overflow-hidden bg-white/45 rounded-[var(--radius-3xl)] border border-white/60 shadow-[0_30px_60px_-15px_rgba(15,23,42,0.08)] backdrop-blur-xl mx-auto"
        style={{ 
          height: 'min(900px, 94dvh)',
          width: 'min(1700px, 98vw)',
          maxWidth: '100%'
        }}
      >
        <AnimatePresence mode="wait">
          {screen === "start" && (
            <motion.div key="start" initial={{ opacity: 0 }} animate={{ opacity: 1 }} exit={{ opacity: 0 }} className="w-full h-full flex items-center justify-center">
              {renderStart()}
            </motion.div>
          )}
          {screen === "config" && (
            <motion.div key="config" initial={{ opacity: 0 }} animate={{ opacity: 1 }} exit={{ opacity: 0 }} className="w-full h-full">
              {renderConfig()}
            </motion.div>
          )}
          {screen === "teams" && (
            <motion.div key="teams" initial={{ opacity: 0 }} animate={{ opacity: 1 }} exit={{ opacity: 0 }} className="w-full h-full">
              {renderTeams()}
            </motion.div>
          )}
          {screen === "game" && (
            <motion.div key="game" initial={{ opacity: 0 }} animate={{ opacity: 1 }} exit={{ opacity: 0 }} className="w-full h-full">
              {renderGame()}
            </motion.div>
          )}
          {screen === "stats" && (
            <motion.div key="stats" initial={{ opacity: 0 }} animate={{ opacity: 1 }} exit={{ opacity: 0 }} className="w-full h-full">
              {renderStats()}
            </motion.div>
          )}
        </AnimatePresence>
      </main>

      {/* 1. REGISTRAR SECCIÓN DIALOG / MODAL */}
      <Dialog open={showSectionModal} onOpenChange={setShowSectionModal}>
        <DialogContent className="sm:max-w-[480px] bg-white border border-slate-200 rounded-3xl p-6 shadow-xl scrollbar-none">
          <DialogHeader className="space-y-1">
            <DialogTitle className="font-black text-slate-900 uppercase italic tracking-tighter" style={{ fontSize: '1.25rem' }}>
              {editingSection ? "Editar Sección de Clase" : "Registrar Nueva Sección"}
            </DialogTitle>
            <DialogDescription className="font-extrabold text-slate-400 text-[10px] uppercase tracking-wider">
              Configura los datos del aula y carga la lista inicial de alumnos
            </DialogDescription>
          </DialogHeader>

          <div className="flex flex-col gap-4 py-4">
            {/* Grado */}
            <div className="flex flex-col gap-1.5">
              <Label className="text-xs font-black uppercase text-slate-700">Grado Escolar / Ciclo</Label>
              <Input 
                placeholder="Ej. Primaria 5° o Secundaria 3°" 
                value={sectionGrado} 
                onChange={(e) => setSectionGrado(e.target.value)}
                className="rounded-xl border-slate-200 focus-visible:ring-indigo-500 font-bold"
              />
            </div>

            {/* Nivel */}
            <div className="flex flex-col gap-1.5">
              <Label className="text-xs font-black uppercase text-slate-700">Nivel Educativo</Label>
              <Select value={sectionNivel} onValueChange={setSectionNivel}>
                <SelectTrigger className="rounded-xl border-slate-200 font-bold">
                  <SelectValue placeholder="Selecciona el nivel" />
                </SelectTrigger>
                <SelectContent className="bg-white border border-slate-200">
                  <SelectItem value="Primaria">Primaria</SelectItem>
                  <SelectItem value="Secundaria">Secundaria</SelectItem>
                  <SelectItem value="Inicial">Inicial</SelectItem>
                  <SelectItem value="Superior / Técnico">Superior / Técnico</SelectItem>
                </SelectContent>
              </Select>
            </div>

            {/* Nombre/Identificador de la sección */}
            <div className="flex flex-col gap-1.5">
              <Label className="text-xs font-black uppercase text-slate-700">Identificador / Nombre de Sección</Label>
              <Input 
                placeholder="Ej. Sección B, Única, San Pedro" 
                value={sectionNombre} 
                onChange={(e) => setSectionNombre(e.target.value)}
                className="rounded-xl border-slate-200 focus-visible:ring-indigo-500 font-bold"
              />
            </div>

            {/* Color Distintivo de la Sección */}
            <div className="flex flex-col gap-1.5">
              <Label className="text-xs font-black uppercase text-slate-700 flex items-center justify-between">
                <span>Color Distintivo del Aula</span>
                <span className="text-[9px] text-indigo-600 font-extrabold uppercase bg-indigo-50/80 px-2 py-0.5 rounded-md">
                  {COLOR_MAPPING[sectionColor]?.title || sectionColor}
                </span>
              </Label>
              <div className="flex items-center gap-2 bg-slate-50 border border-slate-100 p-2.5 rounded-2xl justify-between">
                {Object.keys(COLOR_MAPPING).map((colorKey) => {
                  const info = COLOR_MAPPING[colorKey];
                  const isSelected = sectionColor === colorKey;
                  return (
                    <button
                      key={colorKey}
                      type="button"
                      onClick={() => setSectionColor(colorKey)}
                      className={`w-7 h-7 rounded-full transition-all cursor-pointer flex items-center justify-center relative hover:scale-110 active:scale-95 ${info.bg} ${
                        isSelected 
                          ? 'ring-2 ring-indigo-500 ring-offset-2 scale-105 shadow-sm' 
                          : 'opacity-70 hover:opacity-100'
                      }`}
                      title={info.title}
                    >
                      {isSelected && (
                        <Check className="w-3.5 h-3.5 text-white font-black stroke-[3.5px]" />
                      )}
                    </button>
                  );
                })}
              </div>
            </div>

            {/* Quick Students Loader text area (only for new sections) */}
            {!editingSection && (
              <div className="flex flex-col gap-1.5 mt-1">
                <Label className="text-xs font-black uppercase text-slate-700 flex items-center justify-between">
                  <span>Lista Inicial de Alumnos (Uno por línea)</span>
                  <span className="text-[10px] text-slate-400 font-bold lowercase">opcional</span>
                </Label>
                <textarea 
                  placeholder="Carlos Alberto&#10;María Fernanda&#10;José Luis"
                  value={sectionRawStudentsText}
                  onChange={(e) => setSectionRawStudentsText(e.target.value)}
                  className="w-full h-[100px] p-3 text-sm font-semibold rounded-xl border border-slate-200 focus:outline-none focus:ring-1 focus:ring-indigo-500 bg-slate-50 font-sans"
                />
              </div>
            )}
          </div>

          <DialogFooter className="flex gap-2 sm:justify-end">
            <Button 
              variant="outline" 
              onClick={() => setShowSectionModal(false)}
              className="rounded-xl font-bold uppercase text-xs border-slate-200 text-slate-500 cursor-pointer"
            >
              Cancelar
            </Button>
            <Button 
              onClick={handleCreateOrEditSection}
              className="bg-indigo-600 hover:bg-indigo-700 text-white font-extrabold uppercase tracking-widest text-xs rounded-xl px-5 cursor-pointer shadow-md"
            >
              Guardar Sección
            </Button>
          </DialogFooter>
        </DialogContent>
      </Dialog>


      {/* 2. GESTIONAR ESTUDIANTES DIALOG / MODAL */}
      <Dialog open={showEstudiantesModal} onOpenChange={setShowEstudiantesModal}>
        <DialogContent className="sm:max-w-[480px] bg-white border border-slate-200 rounded-3xl p-6 shadow-xl scrollbar-none flex flex-col max-h-[85vh]">
          <DialogHeader className="space-y-1">
            <DialogTitle className="font-black text-slate-900 uppercase italic tracking-tighter" style={{ fontSize: '1.25rem' }}>
              Alumnos: {activeManageSection?.grado} {activeManageSection?.nombre}
            </DialogTitle>
            <DialogDescription className="font-extrabold text-slate-400 text-[10px] uppercase tracking-wider">
              Permite añadir nuevos estudiantes o eliminar alumnos del roster
            </DialogDescription>
          </DialogHeader>

          {/* New Student Form */}
          <div className="flex gap-2 py-2">
            <Input 
              placeholder="Nombre del nuevo estudiante..." 
              value={newStudentName}
              onChange={(e) => setNewStudentName(e.target.value)}
              onKeyDown={(e) => {
                if (e.key === 'Enter') handleAddStudentToSection();
              }}
              className="rounded-xl border-slate-200 focus-visible:ring-indigo-500 font-bold text-sm h-11"
            />
            <Button 
              onClick={handleAddStudentToSection}
              className="bg-indigo-600 hover:bg-indigo-700 text-white font-extrabold px-4 h-11 rounded-xl shadow-md shrink-0 cursor-pointer"
            >
              <Plus className="w-5 h-5" />
            </Button>
          </div>

          {/* Scrollable Students List */}
          <div className="flex-1 overflow-y-auto no-scrollbar border border-slate-100 rounded-2xl bg-indigo-50/10 p-2 my-2 min-h-[220px] max-h-[350px]">
            {manageStudentsList.length === 0 ? (
              <div className="text-center py-10 flex flex-col items-center justify-center gap-2">
                <Users className="text-slate-300 w-10 h-10" />
                <p className="text-xs font-bold uppercase text-slate-400">Roster Vacío</p>
                <p className="text-[10px] text-slate-400 font-medium">Agrega un estudiante usando el formulario de arriba.</p>
              </div>
            ) : (
              <div className="flex flex-col gap-1.5">
                {manageStudentsList.map((st, i) => (
                  <div key={st.id} className="bg-white border border-slate-100/60 p-2.5 rounded-xl flex items-center justify-between shadow-sm">
                    <div className="flex items-center gap-2 min-w-0">
                      <span className="text-[10px] text-slate-400 font-black tracking-tighter w-4 text-right">{i+1}.</span>
                      <span className="text-xs font-bold text-slate-800 truncate">{st.name}</span>
                    </div>
                    <button 
                      onClick={() => handleDeleteStudentFromSection(st.id)}
                      className="text-slate-300 hover:text-rose-500 p-1.5 rounded-lg hover:bg-rose-50 border border-transparent hover:border-rose-100 transition-all cursor-pointer flex items-center justify-center"
                      title="Eliminar Estudiante"
                    >
                      <Trash2 className="w-3.5 h-3.5" />
                    </button>
                  </div>
                ))}
              </div>
            )}
          </div>

          <DialogFooter className="mt-2 shrink-0">
            <Button 
              onClick={() => setShowEstudiantesModal(false)}
              className="w-full bg-slate-900 hover:bg-slate-800 text-white font-bold uppercase tracking-widest text-xs py-5 rounded-xl cursor-pointer shadow-md"
            >
              Listo / Guardar Cambios
            </Button>
          </DialogFooter>
        </DialogContent>
      </Dialog>

      {/* CONFIRMACIÓN DE ELIMINACIÓN DE SECCIÓN DIALOG / MODAL */}
      <Dialog open={!!sectionIdToDelete} onOpenChange={(open) => { if (!open) setSectionIdToDelete(null); }}>
        <DialogContent className="sm:max-w-[400px] bg-white border border-slate-200 rounded-3xl p-6 shadow-xl scrollbar-none">
          <DialogHeader className="space-y-2">
            <DialogTitle className="font-black text-slate-900 uppercase italic tracking-tighter text-xl">
              ¿Eliminar Sección?
            </DialogTitle>
            <DialogDescription className="font-extrabold text-slate-500 text-xs leading-relaxed uppercase tracking-wider">
              ¿Está seguro/a de que desea eliminar permanentemente esta sección y todos sus estudiantes registrados? Esta acción no se puede deshacer.
            </DialogDescription>
          </DialogHeader>
          <DialogFooter className="flex gap-2 sm:justify-end mt-4">
            <Button 
              variant="outline" 
              onClick={() => setSectionIdToDelete(null)}
              className="rounded-xl font-bold uppercase text-xs border-slate-200 text-slate-500 cursor-pointer h-11"
            >
              Cancelar
            </Button>
            <Button 
              onClick={async () => {
                if (sectionIdToDelete) {
                  const toastId = toast.loading("Eliminando sección de la nube...");
                  try {
                    await deleteSeccionAndStudents(sectionIdToDelete);
                    toast.success("Sección de clase eliminada correctamente.", { id: toastId });
                    if (currentUser) {
                      await loadSeccionesFromDb(currentUser.uid);
                    }
                  } catch (e) {
                    console.error(e);
                    toast.error("Error al eliminar la sección.", { id: toastId });
                  } finally {
                    setSectionIdToDelete(null);
                  }
                }
              }}
              className="bg-rose-600 hover:bg-rose-700 text-white font-extrabold uppercase tracking-widest text-xs rounded-xl px-5 h-11 cursor-pointer shadow-md"
            >
              Sí, Eliminar
            </Button>
          </DialogFooter>
        </DialogContent>
      </Dialog>
    </div>
  );
}

