import React, { useState, useEffect } from "react";
import { 
  ShieldCheck, 
  ShieldAlert, 
  Layers, 
  HardDrive, 
  Search, 
  Lock, 
  LogOut, 
  User, 
  AlertTriangle, 
  RefreshCw, 
  CheckCircle, 
  ExternalLink, 
  ChevronRight, 
  Info,
  Server,
  ToggleLeft,
  ToggleRight,
  Database,
  Brain,
  Sparkles,
  Check,
  Plus,
  Trash2,
  FileText,
  AlertCircle,
  File,
  Folder,
  FileSpreadsheet,
  UploadCloud,
  Sun,
  Moon
} from "lucide-react";
import { motion, AnimatePresence } from "motion/react";
import SharePointUrlGuide from "./components/SharePointUrlGuide";
import DirectoryExplorer from "./components/DirectoryExplorer";
import StructureTable from "./components/StructureTable";
import { Language, translations } from "./translations";

// Mapped Item Type from Backend
interface MappedItem {
  id: string;
  driveId: string;
  name: string;
  library: string;
  type: "Carpeta" | "Archivo";
  size: number;
  lastModified: string;
  customMetadata: Record<string, string>;
  "Estructura actual SharePoint": string;
}

interface UserProfile {
  displayName: string;
  mail: string;
  userPrincipalName: string;
  jobTitle: string;
}

interface ConfigStatus {
  isConfigured: boolean;
  redirectUri: string;
  clientId: string;
  tenantId: string;
  appName: string;
}

export default function App() {
  // Language & Dark mode states
  const [language, setLanguage] = useState<Language>(() => {
    const stored = localStorage.getItem("riea_language");
    return (stored === "es" || stored === "en") ? (stored as Language) : "es";
  });
  
  const [darkMode, setDarkMode] = useState<boolean>(() => {
    const stored = localStorage.getItem("riea_dark_mode");
    if (stored !== null) {
      return stored === "true";
    }
    return window.matchMedia("(prefers-color-scheme: dark)").matches;
  });

  useEffect(() => {
    if (darkMode) {
      document.documentElement.classList.add("dark");
    } else {
      document.documentElement.classList.remove("dark");
    }
    localStorage.setItem("riea_dark_mode", String(darkMode));
  }, [darkMode]);

  useEffect(() => {
    localStorage.setItem("riea_language", language);
  }, [language]);

  const t = translations[language];

  // Config state
  const [config, setConfig] = useState<ConfigStatus>({
    isConfigured: false,
    redirectUri: "",
    clientId: "",
    tenantId: "common",
    appName: "Microsoft Graph & SharePoint Mapper"
  });

  // Auth states
  const [isDemo, setIsDemo] = useState(true); // Default to Demo Mode so it's instantly previewable
  const [accessToken, setAccessToken] = useState<string | null>(null);
  const [user, setUser] = useState<UserProfile | null>(null);
  const [authLoading, setAuthLoading] = useState(false);
  const [authError, setAuthError] = useState<string | null>(null);

  // SharePoint States
  const [siteUrl, setSiteUrl] = useState("https://riearaas.sharepoint.com/sites/Finanzas");
  const [resolvingSite, setResolvingSite] = useState(false);
  const [resolvedSite, setResolvedSite] = useState<any | null>(null);
  const [siteError, setSiteError] = useState<string | null>(null);

  // Structural mapping states
  const [activeTab, setActiveTab] = useState<"explorer" | "mapper">("explorer");
  const [mainTab, setMainTab] = useState<"mapper" | "classifier">("mapper");
  const [mappingLoading, setMappingLoading] = useState(false);
  const [mappedItems, setMappedItems] = useState<MappedItem[]>([]);
  const [mappingError, setMappingError] = useState<string | null>(null);
  const [hasMoreThan15, setHasMoreThan15] = useState(false);
  const [showMapSuccessAlert, setShowMapSuccessAlert] = useState(false);

  // AI Classifier states
  const [selectedFile, setSelectedFile] = useState<MappedItem | null>(null);
  const [classifierSearch, setClassifierSearch] = useState("");
  const [geminiSuggestions, setGeminiSuggestions] = useState<string[]>([]);
  const [editableTags, setEditableTags] = useState<string[]>(["", "", "", "", ""]);
  const [isClassifying, setIsClassifying] = useState(false);
  const [classificationWarning, setClassificationWarning] = useState<string | null>(null);
  const [isApplyingMetadata, setIsApplyingMetadata] = useState(false);
  const [applySuccess, setApplySuccess] = useState<string | null>(null);
  const [applyError, setApplyError] = useState<string | null>(null);

  // Cost-control & Session tracking states
  const [sessionId] = useState(() => {
    let sid = sessionStorage.getItem("riea_session_id");
    if (!sid) {
      sid = "session_" + Math.random().toString(36).substring(2, 15);
      sessionStorage.setItem("riea_session_id", sid);
    }
    return sid;
  });

  const [aiCallCount, setAiCallCount] = useState<number>(() => {
    const stored = sessionStorage.getItem("riea_ai_call_count");
    return stored ? parseInt(stored, 10) : 0;
  });

  const incrementAiCallCount = (newVal?: number) => {
    setAiCallCount((prev) => {
      const next = newVal !== undefined ? newVal : prev + 1;
      sessionStorage.setItem("riea_ai_call_count", String(next));
      return next;
    });
  };

  const [isDragging, setIsDragging] = useState(false);

  const handleLocalFileUpload = (files: FileList | null) => {
    if (!files || files.length === 0) return;
    
    const driveId = resolvedSite?.id || "demo-drive";
    const library = "Cargas Locales";
    
    const newItems: MappedItem[] = Array.from(files).map((file, idx) => {
      const uniqueId = `local-upload-${Date.now()}-${idx}-${Math.random().toString(36).substring(2, 7)}`;
      return {
        id: uniqueId,
        driveId: driveId,
        name: file.name,
        library: library,
        type: "Archivo",
        size: file.size,
        lastModified: new Date(file.lastModified || Date.now()).toISOString(),
        customMetadata: {},
        "Estructura actual SharePoint": `Cargas Locales/${file.name}`
      };
    });
    
    setMappedItems(prev => [...newItems, ...prev]);
    if (newItems.length > 0) {
      setSelectedFile(newItems[0]);
      setGeminiSuggestions([]);
      setClassificationWarning(null);
    }
  };

  // Auto-dismiss mapping success alert after 5 seconds
  useEffect(() => {
    if (showMapSuccessAlert) {
      const timer = setTimeout(() => {
        setShowMapSuccessAlert(false);
      }, 5000);
      return () => clearTimeout(timer);
    }
  }, [showMapSuccessAlert]);

  // Sync editableTags when suggestions or selectedFile changes
  useEffect(() => {
    if (geminiSuggestions.length > 0) {
      setEditableTags(geminiSuggestions);
    } else if (selectedFile) {
      const existing = selectedFile.customMetadata["AI Tags"] || selectedFile.customMetadata.MetaData || "";
      const tags = existing.split(/\s+/).filter(Boolean);
      while (tags.length < 5) {
        tags.push("");
      }
      setEditableTags(tags.slice(0, 5));
    } else {
      setEditableTags(["", "", "", "", ""]);
    }
  }, [geminiSuggestions, selectedFile]);
  
  // Auto-populate site and items in Demo Mode when entering the Classifier tab
  useEffect(() => {
    if (isDemo && mainTab === "classifier") {
      if (!resolvedSite) {
        setResolvedSite({
          id: "site-root-1234",
          displayName: "Contoso SharePoint (Demo Portal)",
          webUrl: "https://riea.sharepoint.com/sites/DemoPortal"
        });
      }
      if (mappedItems.length === 0) {
        const fetchDemoItems = async () => {
          setMappingLoading(true);
          try {
            const response = await fetch("/api/sharepoint/map-structure", {
              method: "POST",
              headers: { "Content-Type": "application/json" },
              body: JSON.stringify({
                siteId: "site-root-1234",
                isDemo: true
              })
            });
            if (response.ok) {
              const data = await response.json();
              const rawItems = data.mappedItems || [];
              if (rawItems.length > 15) {
                setHasMoreThan15(true);
                setMappedItems(rawItems.slice(0, 15));
              } else {
                setHasMoreThan15(false);
                setMappedItems(rawItems);
              }
              setShowMapSuccessAlert(true);
            }
          } catch (e) {
            console.error("Error auto-loading demo items:", e);
          } finally {
            setMappingLoading(false);
          }
        };
        fetchDemoItems();
      }
    }
  }, [isDemo, mainTab, resolvedSite, mappedItems.length]);

  // Load stored credentials and poll/listen for storage changes to handle login callback
  useEffect(() => {
    const checkStorage = () => {
      try {
        const storedToken = localStorage.getItem("microsoft_access_token");
        const storedProfile = localStorage.getItem("microsoft_user_profile");
        if (storedToken && storedProfile) {
          setAccessToken((prev) => {
            if (prev !== storedToken) {
              return storedToken;
            }
            return prev;
          });
          setUser((prev) => {
            const parsed = JSON.parse(storedProfile);
            if (JSON.stringify(prev) !== JSON.stringify(parsed)) {
              return parsed;
            }
            return prev;
          });
          setIsDemo(false); // Force off demo mode upon finding real active credentials
          setAuthError(null);
          setAuthLoading(false);
        }
      } catch (e) {
        console.error("Error reading stored credentials from localStorage:", e);
      }
    };

    // Run immediately on mount
    checkStorage();

    // Listen for storage changes across windows/tabs (same origin)
    window.addEventListener("storage", checkStorage);

    // Poll every 1 second to handle cases where storage events don't propagate across frames/sandboxes
    const interval = setInterval(checkStorage, 1000);

    return () => {
      window.removeEventListener("storage", checkStorage);
      clearInterval(interval);
    };
  }, []);

  // Load backend configuration status
  useEffect(() => {
    const fetchConfig = async () => {
      try {
        const res = await fetch("/api/config");
        if (res.ok) {
          const data = await res.json();
          setConfig(data);
          // If real Microsoft secrets are configured in .env and we have a token, default to Real Mode
          const storedToken = localStorage.getItem("microsoft_access_token");
          if (data.isConfigured && (storedToken || accessToken)) {
            setIsDemo(false);
          }
        }
      } catch (err) {
        console.error("Error loading config:", err);
      }
    };
    fetchConfig();
  }, [accessToken]);

  // Listen for login popup postMessage events
  useEffect(() => {
    const handleOAuthMessage = (event: MessageEvent) => {
      // Security: validate that origin matches preview domain or localhost
      const origin = event.origin;
      if (!origin.endsWith(".run.app") && !origin.includes("localhost") && !origin.includes("127.0.0.1")) {
        return;
      }

      if (event.data?.type === "OAUTH_AUTH_SUCCESS") {
        const { token, user: profile } = event.data;
        if (token) {
          setAccessToken(token);
          setUser(profile);
          setIsDemo(false); // Force off demo mode upon real successful login
          setAuthError(null);
          setAuthLoading(false);
          
          // Clear previous SharePoint states
          setResolvedSite(null);
          setMappedItems([]);
          setSiteError(null);
        }
      }
    };

    window.addEventListener("message", handleOAuthMessage);
    return () => window.removeEventListener("message", handleOAuthMessage);
  }, []);

  // MSAL Login trigger using a separate tab to bypass CSP / sandbox iframe constraints
  const handleMicrosoftLogin = async () => {
    setAuthLoading(true);
    setAuthError(null);
    try {
      const response = await fetch("/api/auth/url");
      if (!response.ok) {
        const errData = await response.json();
        throw new Error(errData.message || "No se pudo obtener la URL de autenticación.");
      }
      
      const { url } = await response.json();
      
      // Open Microsoft OAuth in a new tab/window to bypass CSP / X-Frame-Options/ Sandboxed iframe issues
      const authWindow = window.open(url, "_blank");
      
      if (!authWindow) {
        console.warn("Nueva pestaña bloqueada, reintentando con redirección de ventana principal...");
        // Fallback: If opening a new tab fails or is blocked, break out of the iframe completely to authenticate
        window.top!.location.href = url;
      }
    } catch (err: any) {
      console.error(err);
      setAuthLoading(false);
      setAuthError(err.message || "Ocurrió un error al iniciar sesión.");
    }
  };

  // Resolve SharePoint Site ID & Validate access
  const handleResolveSite = async (e?: React.FormEvent) => {
    if (e) e.preventDefault();
    if (!siteUrl.trim()) return;

    setResolvingSite(true);
    setSiteError(null);
    setResolvedSite(null);
    setMappedItems([]);

    try {
      const headers: Record<string, string> = {
        "Content-Type": "application/json"
      };
      if (!isDemo && accessToken) {
        headers["Authorization"] = `Bearer ${accessToken}`;
      }

      const response = await fetch("/api/sharepoint/resolve-site", {
        method: "POST",
        headers,
        body: JSON.stringify({
          siteUrl: siteUrl.trim(),
          isDemo
        })
      });

      if (!response.ok) {
        const errorData = await response.json();
        
        // Custom message for Access Denied
        if (response.status === 403) {
          throw new Error("Acceso denegado: Su cuenta no cuenta con permisos suficientes (Sites.Read.All) para explorar este sitio de SharePoint, o el sitio no pertenece a su tenant.");
        } else if (response.status === 401) {
          // Token expired handling
          handleLogout();
          throw new Error("Su sesión de Microsoft Graph ha expirado. Por favor, vuelva a iniciar sesión.");
        } else {
          throw new Error(errorData.message || "No se pudo resolver el sitio de SharePoint.");
        }
      }

      const data = await response.json();
      setResolvedSite(data.site);
      
      // Auto-fetch directories
      setActiveTab("explorer");
    } catch (err: any) {
      console.error(err);
      setSiteError(err.message || "Ocurrió un error al validar el acceso al sitio de SharePoint.");
    } finally {
      setResolvingSite(false);
    }
  };

  // Recursive SharePoint structure mapping
  const handleMapCompleteStructure = async () => {
    if (!resolvedSite) return;

    setMappingLoading(true);
    setMappingError(null);
    setMappedItems([]);
    setHasMoreThan15(false);

    try {
      const headers: Record<string, string> = {
        "Content-Type": "application/json"
      };
      if (!isDemo && accessToken) {
        headers["Authorization"] = `Bearer ${accessToken}`;
      }

      const response = await fetch("/api/sharepoint/map-structure", {
        method: "POST",
        headers,
        body: JSON.stringify({
          siteId: resolvedSite.id,
          isDemo
        })
      });

      if (!response.ok) {
        const errData = await response.json();
        throw new Error(errData.message || "Error al mapear la estructura del sitio.");
      }

      const data = await response.json();
      const rawItems = data.mappedItems || [];
      if (rawItems.length > 15) {
        setHasMoreThan15(true);
        setMappedItems(rawItems.slice(0, 15));
      } else {
        setHasMoreThan15(false);
        setMappedItems(rawItems);
      }
      setShowMapSuccessAlert(true);
    } catch (err: any) {
      console.error(err);
      setMappingError(err.message || "Ocurrió un error al generar el mapa de estructura.");
    } finally {
      setMappingLoading(false);
    }
  };

  // Helper for mock demo hashtags based on file/folder names or type (Paso 3)
  const generateDemoHashtags = (fileName: string, type: string): string[] => {
    const nameLower = fileName.toLowerCase();
    const tags: string[] = [];

    if (nameLower.includes("factura") || nameLower.includes("invoice") || nameLower.includes("pago") || nameLower.includes("billing") || nameLower.includes("recibo")) {
      tags.push("#Factura", "#Finanzas", "#Pagos", "#Contabilidad", "#Gastos");
    } else if (nameLower.includes("contrato") || nameLower.includes("contract") || nameLower.includes("legal") || nameLower.includes("acuerdo")) {
      tags.push("#Contrato", "#Legal", "#Acuerdos", "#Corporativo", "#Firmado");
    } else if (nameLower.includes("informe") || nameLower.includes("report") || nameLower.includes("analisis") || nameLower.includes("análisis") || nameLower.includes("resumen")) {
      tags.push("#Informe", "#Reporte", "#Analisis", "#Metricas", "#Documento");
    } else if (nameLower.includes("diseño") || nameLower.includes("design") || nameLower.includes("imagen") || nameLower.includes("logo") || nameLower.includes("png") || nameLower.includes("jpg")) {
      tags.push("#Diseno", "#Multimedia", "#Marketing", "#Brand", "#Assets");
    } else if (nameLower.includes("presupuesto") || nameLower.includes("budget") || nameLower.includes("forecast") || nameLower.includes("estimación")) {
      tags.push("#Presupuesto", "#Planificacion", "#Finanzas", "#Forecast", "#Costos");
    } else if (nameLower.includes("curriculum") || nameLower.includes("cv") || nameLower.includes("empleado") || nameLower.includes("nomina") || nameLower.includes("rrhh") || nameLower.includes("hr")) {
      tags.push("#RRHH", "#Personal", "#Talento", "#Contratacion", "#Nomina");
    } else if (type === "Carpeta") {
      tags.push("#Carpeta", "#Organizacion", "#Directorio", "#Compartido", "#Estructura");
    } else {
      tags.push("#Documento", "#General", "#SharePoint", "#Clasificado", "#Archivo");
    }

    // Ensure we always return exactly 5 tags
    while (tags.length < 5) {
      tags.push("#General");
    }
    return tags.slice(0, 5);
  };

  // AI Classification suggesting categories using Gemini
  const handleSuggestCategories = async (file: MappedItem) => {
    if (aiCallCount >= 100) {
      setApplyError("Límite de procesamiento de IA alcanzado por seguridad del servidor. Si necesitas procesar volúmenes masivos, contacta a RIEA para una solución dedicada corporativa.");
      return;
    }

    setIsClassifying(true);
    setClassificationWarning(null);
    setApplySuccess(null);
    setApplyError(null);
    setGeminiSuggestions([]);

    if (isDemo) {
      // MODO DEMO: Simular carga de 1.5 segundos y generar hashtags falsos pero coherentes
      await new Promise((resolve) => setTimeout(resolve, 1500));
      const demoTags = generateDemoHashtags(file.name, file.type);
      setGeminiSuggestions(demoTags);
      setIsClassifying(false);
      incrementAiCallCount();
      return;
    }

    try {
      const response = await fetch("/api/gemini/suggest-categories", {
        method: "POST",
        headers: { 
          "Content-Type": "application/json",
          "x-session-id": sessionId
        },
        body: JSON.stringify({
          fileName: file.name,
          library: file.library,
          type: file.type,
          customMetadata: file.customMetadata,
          path: file["Estructura actual SharePoint"]
        })
      });

      if (!response.ok) {
        if (response.status === 429) {
          const errData = await response.json();
          incrementAiCallCount(100); // Force sync count to 100
          throw new Error(errData.message || "Límite de procesamiento de IA alcanzado por seguridad del servidor.");
        }
        throw new Error("La llamada al servidor de clasificación falló.");
      }

      const data = await response.json();
      setGeminiSuggestions(data.hashtags || []);
      
      if (data.sessionCount !== undefined) {
        incrementAiCallCount(data.sessionCount);
      } else {
        incrementAiCallCount();
      }

      if (data.warning) {
        setClassificationWarning(data.warning);
      }
    } catch (err: any) {
      console.error(err);
      setApplyError(err.message || "No se pudieron obtener sugerencias de clasificación.");
    } finally {
      setIsClassifying(false);
    }
  };

  // Helper function to export classification reports in CSV or JSON
  const handleExport = (format: "csv" | "json") => {
    if (mappedItems.length === 0) return;

    if (format === "csv") {
      const headers = ["Nombre de Elemento", "Tipo", "Biblioteca", "Ruta SharePoint", "AI Tags"];
      
      const rows = mappedItems.map(item => {
        const name = (item.name || "").replace(/"/g, '""');
        const type = (item.type || "").replace(/"/g, '""');
        const library = (item.library || "").replace(/"/g, '""');
        const path = (item["Estructura actual SharePoint"] || "").replace(/"/g, '""');
        const aiTags = (item.customMetadata?.["AI Tags"] || item.customMetadata?.MetaData || "").replace(/"/g, '""');
        
        return `"${name}","${type}","${library}","${path}","${aiTags}"`;
      });
      
      // UTF-8 BOM to make sure accents like 'ó' display nicely in Excel in Spanish
      const csvContent = "\uFEFF" + [headers.join(","), ...rows].join("\n");
      const blob = new Blob([csvContent], { type: "text/csv;charset=utf-8;" });
      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.setAttribute("href", url);
      link.setAttribute("download", `Reporte_Clasificacion_RIEA_${new Date().toISOString().split("T")[0]}.csv`);
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
    } else {
      const jsonContent = JSON.stringify(mappedItems, null, 2);
      const blob = new Blob([jsonContent], { type: "application/json" });
      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.setAttribute("href", url);
      link.setAttribute("download", `Reporte_Clasificacion_RIEA_${new Date().toISOString().split("T")[0]}.json`);
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
    }
  };

  // Applying final tags to the SharePoint 'AI Tags' column
  const handleApplyMetadata = async (tagString: string) => {
    if (!selectedFile) return;
    setIsApplyingMetadata(true);
    setApplySuccess(null);
    setApplyError(null);

    if (isDemo) {
      // MODO DEMO: Simular éxito en pantalla sin tocar SharePoint
      try {
        await new Promise((resolve) => setTimeout(resolve, 1000));
        setApplySuccess("¡Metadatos (Modo Demo) actualizados con éxito en la simulación!");
        
        // Update our local state mappedItems so the table and side-list show the new metadata!
        setMappedItems((prev) =>
          prev.map((item) => {
            if (item.id === selectedFile.id) {
              return {
                ...item,
                customMetadata: {
                  ...item.customMetadata,
                  "AI Tags": tagString
                }
              };
            }
            return item;
          })
        );

        // Also update selectedFile state
        setSelectedFile((prev) => {
          if (!prev) return null;
          return {
            ...prev,
            customMetadata: {
              ...prev.customMetadata,
              "AI Tags": tagString
            }
          };
        });
      } catch (err: any) {
        setApplyError("Error simulando la actualización.");
      } finally {
        setIsApplyingMetadata(false);
      }
      return;
    }

    try {
      const headers: Record<string, string> = {
        "Content-Type": "application/json"
      };
      if (!isDemo && accessToken) {
        headers["Authorization"] = `Bearer ${accessToken}`;
      }

      const response = await fetch("/api/sharepoint/update-metadata", {
        method: "POST",
        headers,
        body: JSON.stringify({
          driveId: selectedFile.driveId,
          itemId: selectedFile.id,
          metadataValue: tagString,
          isDemo
        })
      });

      if (!response.ok) {
        const errorData = await response.json();
        throw new Error(errorData.message || "Error al actualizar la columna 'AI Tags' en SharePoint.");
      }

      const data = await response.json();
      setApplySuccess(data.message || "¡Metadatos actualizados con éxito!");
      
      // Update our local state mappedItems so the table and side-list show the new metadata!
      setMappedItems((prev) =>
        prev.map((item) => {
          if (item.id === selectedFile.id) {
            return {
              ...item,
              customMetadata: {
                ...item.customMetadata,
                "AI Tags": tagString
              }
            };
          }
          return item;
        })
      );

      // Also update selectedFile state
      setSelectedFile((prev) => {
        if (!prev) return null;
        return {
          ...prev,
          customMetadata: {
            ...prev.customMetadata,
            "AI Tags": tagString
          }
        };
      });

    } catch (err: any) {
      console.error(err);
      setApplyError(err.message || "Error al intentar actualizar la columna 'AI Tags' en SharePoint.");
    } finally {
      setIsApplyingMetadata(false);
    }
  };

  const handleLogout = () => {
    setAccessToken(null);
    setUser(null);
    setResolvedSite(null);
    setMappedItems([]);
    setSiteError(null);
    setAuthError(null);
    try {
      localStorage.removeItem("microsoft_access_token");
      localStorage.removeItem("microsoft_user_profile");
    } catch (e) {
      console.error("Error clearing stored credentials on logout:", e);
    }
  };

  const toggleMode = () => {
    // Reset all session/connected states when switching modes
    handleLogout();
    setIsDemo(!isDemo);
  };

  return (
    <div id="main-layout" className="w-screen h-screen overflow-hidden bg-[#F8FAFC] dark:bg-slate-900 font-sans text-slate-900 dark:text-slate-100 antialiased flex flex-col transition-colors duration-300">
      {/* HEADER BAR */}
      <header id="app-header" className="bg-white dark:bg-slate-800 border-b border-slate-200 dark:border-slate-700 sticky top-0 z-40 px-6 py-4 shrink-0 transition-colors duration-300">
        <div className="max-w-7xl mx-auto flex flex-col md:flex-row md:items-center justify-between gap-4">
          
          {/* Logo & title */}
          <div className="flex items-center gap-4">
            {/* RIEA STYLIZED LOGO */}
            <div className="flex items-center gap-1.5 shrink-0 bg-slate-950 px-3.5 py-1.5 rounded-lg text-white font-serif tracking-wider font-extrabold text-sm shadow-sm select-none border border-slate-800">
              <span className="text-white italic">R</span>
              <span className="text-white italic">I</span>
              <span className="text-white italic">E</span>
              <span className="text-white italic">A</span>
            </div>
            
            <div className="h-8 w-[1px] bg-slate-200 dark:bg-slate-700 hidden sm:block"></div>

            <div>
              <div className="flex items-center gap-2">
                <h1 className="text-lg font-semibold tracking-tight text-slate-800 dark:text-slate-100 leading-none">
                  {t.appTitle}
                </h1>
              </div>
              <p className="text-xs text-slate-500 dark:text-slate-400 mt-0.5 font-normal leading-normal">
                {t.appSubtitle}
              </p>
            </div>
          </div>

          {/* Mode Toggler, Lang Selector, Theme Toggle and User Card */}
          <div className="flex flex-wrap items-center gap-3 md:gap-4">
            
            {/* Language Selector Dropdown */}
            <div className="flex items-center gap-1.5 bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 px-2.5 py-1.5 rounded-lg text-xs font-medium text-slate-700 dark:text-slate-250 shadow-xs">
              <span className="text-sm select-none leading-none">
                {language === "es" ? "🇪🇸" : "🇺🇸"}
              </span>
              <select
                id="language-selector"
                value={language}
                onChange={(e) => setLanguage(e.target.value as Language)}
                className="bg-transparent border-none text-xs font-semibold focus:outline-none focus:ring-0 cursor-pointer text-slate-700 dark:text-slate-200"
              >
                <option value="es" className="dark:bg-slate-800">ES</option>
                <option value="en" className="dark:bg-slate-800">EN</option>
              </select>
            </div>

            {/* Dark Mode Toggle Button */}
            <button
              id="theme-toggle-btn"
              onClick={() => setDarkMode(!darkMode)}
              className="p-2 bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded-lg text-slate-500 dark:text-slate-400 hover:bg-slate-50 dark:hover:bg-slate-700 transition-all shadow-xs cursor-pointer flex items-center justify-center shrink-0"
              title={darkMode ? "Switch to Light Mode" : "Switch to Dark Mode"}
            >
              {darkMode ? (
                <Sun className="w-4 h-4 text-yellow-500" />
              ) : (
                <Moon className="w-4 h-4 text-slate-500" />
              )}
            </button>

            {/* Mode Switcher */}
            <div className="flex items-center gap-2 bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 px-3 py-1.5 rounded-lg text-xs font-medium transition-colors">
              <span className={`transition-colors ${isDemo ? "text-amber-600 dark:text-amber-400 font-semibold" : "text-slate-400 dark:text-slate-500"}`}>
                {t.demoMode}
              </span>
              <button 
                id="mode-toggle-btn"
                onClick={toggleMode} 
                className="focus:outline-none cursor-pointer"
                title={isDemo ? (language === "es" ? "Cambiar a Modo Real" : "Switch to Real Mode") : (language === "es" ? "Cambiar a Modo Demo" : "Switch to Demo Mode")}
              >
                {isDemo ? (
                  <ToggleLeft className="w-8 h-8 text-amber-500 hover:text-amber-600 dark:text-amber-400" />
                ) : (
                  <ToggleRight className="w-8 h-8 text-blue-600 hover:text-blue-700 dark:text-blue-400" />
                )}
              </button>
              <span className={`transition-colors ${!isDemo ? "text-blue-600 dark:text-blue-400 font-semibold" : "text-slate-400 dark:text-slate-500"}`}>
                {t.realMode}
              </span>
            </div>

            {/* User Profile Card */}
            {user ? (
              <div className="flex items-center gap-2.5 bg-slate-50 dark:bg-slate-900 border border-slate-200 dark:border-slate-700 px-3 py-1.5 rounded-lg text-xs">
                <div className="w-8 h-8 rounded-full bg-slate-100 dark:bg-slate-800 border border-slate-200 dark:border-slate-700 flex items-center justify-center text-slate-500 dark:text-slate-400 font-medium shrink-0">
                  {user.displayName.split(" ").map(n => n[0]).join("").substring(0, 2).toUpperCase()}
                </div>
                <div className="text-left leading-none max-w-[180px] truncate">
                  <p className="font-medium text-slate-800 dark:text-slate-200 truncate">{user.mail}</p>
                  <p className="text-[10px] uppercase tracking-wider text-green-600 dark:text-green-400 font-bold mt-0.5">{t.authActive}</p>
                </div>
                <button
                  id="logout-btn"
                  onClick={handleLogout}
                  className="p-1.5 hover:bg-slate-100 dark:hover:bg-slate-750 rounded text-slate-500 dark:text-slate-400 transition-colors cursor-pointer"
                  title={t.logout}
                >
                  <LogOut className="w-4 h-4" />
                </button>
              </div>
            ) : !isDemo ? (
              <span className="text-xs text-rose-600 dark:text-rose-400 font-semibold bg-rose-50 dark:bg-rose-950/20 border border-rose-200 dark:border-rose-900/30 px-3 py-1.5 rounded-lg flex items-center gap-1.5">
                <Lock className="w-3.5 h-3.5" />
                {t.notAuthenticated}
              </span>
            ) : (
              <span className="text-xs text-amber-600 dark:text-amber-400 font-semibold bg-amber-50 dark:bg-amber-950/20 border border-amber-200 dark:border-amber-900/30 px-3 py-1.5 rounded-lg flex items-center gap-1.5">
                <User className="w-3.5 h-3.5" />
                {t.demoSession}
              </span>
            )}

          </div>
        </div>
      </header>

      {/* BODY CONTENT CONTAINER */}
      <main id="app-body" className="flex-1 max-w-7xl w-full mx-auto p-6 space-y-6 overflow-y-auto overflow-x-hidden">
        
        {/* MAIN NAVIGATION TABS */}
        <div id="main-navigation-tabs" className="flex border-b border-slate-200 dark:border-slate-700 bg-white dark:bg-slate-800 p-1 rounded-xl border max-w-md shadow-xs mx-auto md:mx-0 transition-colors">
          <button
            id="main-tab-mapper"
            onClick={() => setMainTab("mapper")}
            className={`flex-1 py-2.5 text-xs font-semibold rounded-lg transition-all cursor-pointer flex items-center justify-center gap-2 ${
              mainTab === "mapper"
                ? "bg-slate-900 dark:bg-slate-700 text-white shadow-sm font-bold"
                : "text-slate-500 dark:text-slate-400 hover:text-slate-800 dark:hover:text-slate-200 hover:bg-slate-50 dark:hover:bg-slate-750"
            }`}
          >
            <Layers className="w-4 h-4" />
            {t.tabMapper}
          </button>
          <button
            id="main-tab-classifier"
            onClick={() => setMainTab("classifier")}
            className={`flex-1 py-2.5 text-xs font-semibold rounded-lg transition-all cursor-pointer flex items-center justify-center gap-2 ${
              mainTab === "classifier"
                ? "bg-slate-900 dark:bg-slate-700 text-white shadow-sm font-bold"
                : "text-slate-500 dark:text-slate-400 hover:text-slate-800 dark:hover:text-slate-200 hover:bg-slate-50 dark:hover:bg-slate-750"
            }`}
          >
            <Brain className="w-4 h-4 text-purple-500" />
            {t.tabClassifier}
          </button>
        </div>

        {/* Warning If Demo Mode is Active */}
        <AnimatePresence>
          {isDemo && (
            <motion.div
              initial={{ opacity: 0, y: -10 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -10 }}
              className="bg-amber-50/50 dark:bg-amber-950/20 border border-amber-200 dark:border-amber-900/30 p-4 rounded-xl flex flex-col sm:flex-row items-start sm:items-center justify-between gap-4 shadow-sm"
            >
              <div className="flex gap-3">
                <div className="p-2 bg-amber-100/60 dark:bg-amber-900/40 text-amber-800 dark:text-amber-300 rounded-lg shrink-0 mt-0.5 sm:mt-0">
                  <AlertTriangle className="w-5 h-5" />
                </div>
                <div>
                  <h4 className="font-semibold text-amber-900 dark:text-amber-200 text-sm">{t.demoWarningTitle}</h4>
                  <p className="text-xs text-amber-700 dark:text-amber-400 mt-0.5 leading-relaxed font-normal">
                    {t.demoWarningText}
                  </p>
                </div>
              </div>
              <button
                onClick={() => setIsDemo(false)}
                className="text-xs font-semibold text-amber-900 dark:text-amber-300 bg-amber-100/60 dark:bg-amber-900/40 hover:bg-amber-200/80 dark:hover:bg-amber-900/60 px-3.5 py-1.5 rounded-lg border border-amber-200 dark:border-amber-800/40 whitespace-nowrap cursor-pointer transition-colors"
              >
                {t.demoWarningButton}
              </button>
            </motion.div>
          )}
        </AnimatePresence>

        {/* MAIN VIEWPORT SWITCHER */}
        {mainTab === "mapper" ? (
          <div id="view-mapper" className="space-y-6">
            {/* TOP CONFIGURATION / LOGIN ROW */}
            <div className="grid grid-cols-1 lg:grid-cols-12 gap-6">
              
              {/* LEFT CARD: AUTHENTICATION AND CONNECT SITE */}
              <div className="lg:col-span-7 bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded-xl p-6 shadow-sm flex flex-col justify-between space-y-5 transition-colors">
                <div>
                  <h3 className="font-semibold text-slate-800 dark:text-slate-100 text-base flex items-center gap-2">
                    <ShieldCheck className="w-5 h-5 text-blue-600 dark:text-blue-400" />
                    {language === "es" ? "1. Conexión de Seguridad (Azure AD & SharePoint)" : "1. Security Connection (Azure AD & SharePoint)"}
                  </h3>
                  <p className="text-xs text-slate-500 dark:text-slate-400 mt-1">
                    {language === "es" ? "Autentícate de forma segura utilizando OAuth2 y luego resuelve la ruta de tu SharePoint Site." : "Securely authenticate using OAuth2 and then resolve the path of your SharePoint Site."}
                  </p>
                </div>

                {/* REAL MODE LOGIN CARD */}
                {!isDemo && (
                  <div className="p-4 rounded-lg border border-blue-100 dark:border-blue-900/30 bg-blue-50/20 dark:bg-blue-950/10 space-y-3.5">
                    <div className="flex items-center justify-between">
                      <span className="text-xs font-semibold text-blue-950 dark:text-blue-300 uppercase tracking-wide">{t.stepATitle}</span>
                      {accessToken ? (
                        <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full bg-emerald-100 dark:bg-emerald-950/40 text-emerald-800 dark:text-emerald-400 text-[10px] font-bold">
                          <CheckCircle className="w-3 h-3" /> {t.stepAConnected}
                        </span>
                      ) : (
                        <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full bg-slate-100 dark:bg-slate-700 text-slate-600 dark:text-slate-300 text-[10px] font-bold">
                          {t.stepARequired}
                        </span>
                      )}
                    </div>

                    {!accessToken ? (
                      <div className="space-y-3">
                        <button
                          id="microsoft-login-btn"
                          onClick={handleMicrosoftLogin}
                          disabled={authLoading}
                          className="w-full flex items-center justify-center gap-2.5 px-4 py-3 bg-blue-600 hover:bg-blue-700 disabled:bg-blue-300 text-white font-semibold text-xs rounded-lg shadow-sm transition-all duration-200 cursor-pointer"
                        >
                          {authLoading ? (
                            <>
                              <RefreshCw className="w-4 h-4 animate-spin" />
                              <span>{language === "es" ? "Obteniendo URL de MSAL..." : "Obtaining MSAL URL..."}</span>
                            </>
                          ) : (
                            <>
                              <svg className="w-4 h-4" viewBox="0 0 23 23" fill="none" xmlns="http://www.w3.org/2000/svg">
                                <path d="M0 0H11V11H0V0Z" fill="#F25022"/>
                                <path d="M12 0H23V11H12V0Z" fill="#7FBA00"/>
                                <path d="M0 12H11V23H0V12Z" fill="#00A4EF"/>
                                <path d="M12 12H23V23H12V12Z" fill="#FFB900"/>
                              </svg>
                              <span>{t.stepAButton}</span>
                            </>
                          )}
                        </button>
                        {authError && (
                          <div className="p-3 bg-rose-50 dark:bg-rose-950/20 text-rose-700 dark:text-rose-400 text-xs rounded-lg border border-rose-100 dark:border-rose-900/30 font-medium">
                            {authError}
                          </div>
                        )}
                      </div>
                    ) : (
                      <div className="flex items-center justify-between p-3 bg-white dark:bg-slate-900 border border-blue-100 dark:border-blue-900/30 rounded-lg text-xs">
                        <div>
                          <p className="font-semibold text-slate-800 dark:text-slate-200">{user?.displayName}</p>
                          <p className="text-slate-500 dark:text-slate-400 font-mono text-[10px] mt-0.5">{user?.mail}</p>
                        </div>
                        <span className="text-[10px] font-medium bg-slate-100 dark:bg-slate-700 px-2 py-1 rounded text-slate-500 dark:text-slate-350">
                          {language === "es" ? "Token obtenido" : "Token obtained"}
                        </span>
                      </div>
                    )}
                  </div>
                )}

                {/* CONNECT SHAREPOINT SITE FORM */}
                <form onSubmit={handleResolveSite} className="p-4 rounded-lg border border-slate-200 dark:border-slate-700 bg-slate-50/50 dark:bg-slate-900/30 space-y-3 transition-colors">
                  <div className="flex items-center justify-between">
                    <span className="text-xs font-semibold text-slate-700 dark:text-slate-300 uppercase tracking-wide">
                      {isDemo ? (language === "es" ? "Selecciona o busca un Sitio de SharePoint" : "Select or search a SharePoint Site") : t.stepBTitle}
                    </span>
                    {resolvedSite ? (
                      <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full bg-emerald-100 dark:bg-emerald-950/40 text-emerald-800 dark:text-emerald-400 text-[10px] font-bold">
                        <CheckCircle className="w-3 h-3" /> {t.stepBReady}
                      </span>
                    ) : (
                      <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full bg-amber-100 dark:bg-amber-950/40 text-amber-800 dark:text-amber-400 text-[10px] font-bold">
                        {t.stepBPending}
                      </span>
                    )}
                  </div>

                  <div className="space-y-1">
                    <label className="text-[11px] font-bold text-slate-500 dark:text-slate-400 block">{t.stepBLabel}</label>
                    <div className="flex flex-col sm:flex-row gap-2">
                      <input
                        type="text"
                        value={siteUrl}
                        onChange={(e) => setSiteUrl(e.target.value)}
                        placeholder={t.stepBPlaceholder}
                        className="flex-1 text-xs border border-slate-200 dark:border-slate-700 rounded-lg px-3.5 py-2.5 focus:outline-none focus:ring-2 focus:ring-blue-500/10 focus:border-blue-500 bg-white dark:bg-slate-800 dark:text-slate-100"
                      />
                      <button
                        id="connect-site-btn"
                        type="submit"
                        disabled={resolvingSite || (!isDemo && !accessToken)}
                        className="flex items-center justify-center gap-1.5 px-4 py-2.5 bg-slate-900 dark:bg-slate-700 hover:bg-slate-800 dark:hover:bg-slate-600 disabled:bg-slate-200 dark:disabled:bg-slate-800 text-white disabled:text-slate-400 dark:disabled:text-slate-500 text-xs font-semibold rounded-lg transition-colors shrink-0 cursor-pointer"
                      >
                        {resolvingSite ? (
                          <>
                            <RefreshCw className="w-3.5 h-3.5 animate-spin" />
                            <span>{language === "es" ? "Validando..." : "Validating..."}</span>
                          </>
                        ) : (
                          <>
                            <Database className="w-3.5 h-3.5" />
                            <span>{language === "es" ? "Validar Acceso" : "Validate Access"}</span>
                          </>
                        )}
                      </button>
                    </div>
                    {!isDemo && !accessToken && (
                      <span className="text-[10px] text-indigo-600 dark:text-indigo-400 font-medium block mt-1">
                        {language === "es" ? "* Debes iniciar sesión con Microsoft (Paso A) antes de validar acceso." : "* You must sign in with Microsoft (Step A) before validating access."}
                      </span>
                    )}
                  </div>
                </form>
              </div>

              {/* RIGHT CARD: CONFIG STATUS / SUMMARY GUIDE OR ERROR PANEL */}
              <div className="lg:col-span-5 bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded-xl p-6 shadow-sm flex flex-col justify-between transition-colors">
                <div className="space-y-3.5">
                  <h3 className="font-semibold text-slate-800 dark:text-slate-100 text-base flex items-center gap-2">
                    <Info className="w-5 h-5 text-blue-600 dark:text-blue-400" />
                    {language === "es" ? "Estado del Servidor & Credenciales" : "Server Status & Credentials"}
                  </h3>
                  
                  <div className="space-y-2.5">
                    {/* Credentials active */}
                    <div className="flex items-center justify-between p-3 bg-[#F8FAFC] dark:bg-slate-900/50 rounded-lg border border-slate-200 dark:border-slate-700 text-xs font-medium transition-colors">
                      <span className="text-slate-500 dark:text-slate-400">{language === "es" ? "Credenciales Azure (.env):" : "Azure Credentials (.env):"}</span>
                      {config.isConfigured ? (
                        <span className="text-emerald-700 bg-emerald-100/80 border border-emerald-200 dark:text-emerald-400 dark:bg-emerald-950/40 dark:border-emerald-900/40 px-2.5 py-0.5 rounded-full text-[10px] font-bold">
                          {language === "es" ? "Configuradas" : "Configured"}
                        </span>
                      ) : (
                        <span className="text-amber-700 bg-amber-100/80 border border-amber-200 dark:text-amber-400 dark:bg-amber-950/40 dark:border-amber-900/40 px-2.5 py-0.5 rounded-full text-[10px] font-bold">
                          {language === "es" ? "Ausentes (Modo Demo Activo)" : "Absent (Demo Mode Active)"}
                        </span>
                      )}
                    </div>

                    {/* Redirect URI active */}
                    <div className="flex items-center justify-between p-3 bg-[#F8FAFC] dark:bg-slate-900/50 rounded-lg border border-slate-200 dark:border-slate-700 text-xs font-medium transition-colors">
                      <span className="text-slate-500 dark:text-slate-400">{language === "es" ? "Tenant Microsoft AD:" : "Microsoft AD Tenant:"}</span>
                      <code className="text-slate-600 dark:text-slate-300 font-mono text-[10px] bg-slate-100 dark:bg-slate-800 px-1.5 py-0.5 rounded">{config.tenantId}</code>
                    </div>

                    {/* Scope status */}
                    <div className="p-3 bg-[#F8FAFC] dark:bg-slate-900/50 rounded-lg border border-slate-200 dark:border-slate-700 text-xs font-medium space-y-1.5 transition-colors">
                      <span className="text-slate-500 dark:text-slate-400 block">{language === "es" ? "Scopes Solicitados en MSAL:" : "Requested MSAL Scopes:"}</span>
                      <div className="flex flex-wrap gap-1">
                        {["User.Read", "Sites.Read.All", "Files.Read.All"].map(s => (
                          <span key={s} className="px-2 py-1 bg-slate-50 dark:bg-slate-800 border border-slate-200 dark:border-slate-700 text-slate-600 dark:text-slate-350 text-[10px] font-mono rounded">
                            {s}
                          </span>
                        ))}
                      </div>
                    </div>
                  </div>
                </div>

                {/* SECURITY WARNING / GRACEFUL EXCEPTION PANEL */}
                <div className="mt-4">
                  <AnimatePresence mode="wait">
                    {siteError ? (
                      <motion.div
                        key="site-error"
                        initial={{ opacity: 0, scale: 0.95 }}
                        animate={{ opacity: 1, scale: 1 }}
                        exit={{ opacity: 0, scale: 0.95 }}
                        className="p-4 bg-rose-50 dark:bg-rose-950/20 border border-rose-200 dark:border-rose-900/30 rounded-xl text-xs space-y-2.5 transition-colors"
                      >
                        <div className="flex items-center gap-2 text-rose-800 dark:text-rose-400 font-bold uppercase tracking-wider text-[10px]">
                          <ShieldAlert className="w-4 h-4 text-rose-600" />
                          {language === "es" ? "Acceso Denegado / Error de Validación" : "Access Denied / Validation Error"}
                        </div>
                        <p className="text-rose-700 dark:text-rose-300 font-medium leading-relaxed">
                          {siteError}
                        </p>
                        <div className="text-[10px] text-rose-600/90 dark:text-rose-400 leading-relaxed border-t border-rose-100 dark:border-rose-900/20 pt-2 bg-rose-50/50 dark:bg-rose-950/10 transition-colors">
                          <strong>{language === "es" ? "Causa común:" : "Common cause:"}</strong> {language === "es" ? "Tu App Registration en Azure AD requiere permisos delegados autorizados de Sites.Read.All y Files.Read.All, y tu cuenta de Office 365 debe tener permisos de acceso a ese sitio web." : "Your App Registration in Azure AD requires authorized delegated permissions of Sites.Read.All and Files.Read.All, and your Office 365 account must have access permissions to that website."}
                        </div>
                      </motion.div>
                    ) : resolvedSite ? (
                      <motion.div
                        key="site-success"
                        initial={{ opacity: 0, scale: 0.95 }}
                        animate={{ opacity: 1, scale: 1 }}
                        exit={{ opacity: 0, scale: 0.95 }}
                        className="p-4 bg-emerald-50 dark:bg-emerald-950/20 border border-emerald-200 dark:border-emerald-900/30 rounded-xl text-xs space-y-2 transition-colors"
                      >
                        <div className="flex items-center gap-1.5 text-emerald-800 dark:text-emerald-400 font-bold uppercase tracking-wider text-[10px]">
                          <CheckCircle className="w-4 h-4 text-emerald-600" />
                          {t.accessValidated}
                        </div>
                        <p className="text-slate-700 dark:text-slate-300">
                          {t.verifiedSiteText} <strong className="text-slate-950 dark:text-slate-100">{resolvedSite.displayName}</strong>.
                        </p>
                        <a
                          href={resolvedSite.webUrl}
                          target="_blank"
                          rel="noopener noreferrer"
                          className="inline-flex items-center gap-1 text-emerald-700 dark:text-emerald-400 hover:underline font-bold text-[10px] mt-1"
                        >
                          {language === "es" ? "Abrir SharePoint original" : "Open original SharePoint"}
                          <ExternalLink className="w-3 h-3" />
                        </a>
                      </motion.div>
                    ) : (
                      <motion.div
                        key="site-pending"
                        initial={{ opacity: 0, scale: 0.95 }}
                        animate={{ opacity: 1, scale: 1 }}
                        exit={{ opacity: 0, scale: 0.95 }}
                        className="p-4 bg-slate-50 dark:bg-slate-900/40 border border-slate-200 dark:border-slate-700 rounded-xl text-xs flex items-center gap-2.5 text-slate-500 dark:text-slate-400 font-medium transition-colors"
                      >
                        <Server className="w-5 h-5 text-slate-400 shrink-0" />
                        <span>{language === "es" ? "Esperando validación de acceso al sitio de SharePoint. Introduce una URL o mantén el Modo Demo para ver los datos simulados." : "Waiting for SharePoint site access validation. Enter a URL or keep Demo Mode to see simulated data."}</span>
                      </motion.div>
                    )}
                  </AnimatePresence>
                </div>
              </div>

            </div>

            {/* DYNAMIC URL GUIDE REPLACING THE AZURE AD PORTAL GUIDE */}
            <SharePointUrlGuide language={language} />

            {/* ACTIVE EXPLORER & MAPPER DASHBOARD */}
            {resolvedSite ? (
              <div id="active-dashboard" className="space-y-6">
                
                {/* Nav tabs bar */}
                <div className="flex border-b border-slate-200 dark:border-slate-700 bg-white dark:bg-slate-800 p-1 rounded-lg border max-w-md shadow-xs transition-colors">
                  <button
                    id="explorer-tab"
                    onClick={() => setActiveTab("explorer")}
                    className={`flex-1 py-2 text-xs font-semibold rounded transition-all cursor-pointer ${
                      activeTab === "explorer"
                        ? "bg-blue-600 text-white shadow-sm"
                        : "text-slate-500 dark:text-slate-400 hover:text-slate-800 dark:hover:text-slate-200"
                    }`}
                  >
                    {t.subTabExplorer}
                  </button>
                  <button
                    id="mapper-tab"
                    onClick={() => {
                      setActiveTab("mapper");
                      // Auto load complete mapper data if not loaded yet
                      if (mappedItems.length === 0) {
                        handleMapCompleteStructure();
                      }
                    }}
                    className={`flex-1 py-2 text-xs font-semibold rounded transition-all cursor-pointer ${
                      activeTab === "mapper"
                        ? "bg-blue-600 text-white shadow-sm"
                        : "text-slate-500 hover:text-slate-800"
                    }`}
                  >
                    {language === "es" ? "2. Mapa de Estructura Completa" : "2. Complete Structure Map"}
                  </button>
                </div>
 
                {/* TAB 1: INTERACTIVE LAZY EXPLORER */}
                {activeTab === "explorer" && (
                  <motion.div
                    initial={{ opacity: 0, y: 15 }}
                    animate={{ opacity: 1, y: 0 }}
                    transition={{ duration: 0.3 }}
                    className="space-y-4"
                  >
                    <div className="bg-blue-50/50 dark:bg-blue-950/10 border border-blue-100 dark:border-blue-900/30 p-4 rounded-lg flex items-start gap-3 shadow-sm transition-colors">
                      <Info className="w-5 h-5 text-blue-600 dark:text-blue-400 shrink-0 mt-0.5" />
                      <div className="text-xs text-blue-950 dark:text-blue-300 font-medium">
                        <p className="font-semibold">{language === "es" ? "Visualizador Interactivo de Directorios" : "Interactive Directory Viewer"}</p>
                        <p className="mt-0.5 leading-relaxed text-blue-900/90 dark:text-blue-400 font-normal">
                          {language === "es" ? "Navega en tiempo real a través de las bibliotecas de documentos (Drives) de SharePoint. Este explorador consulta dinámicamente la API de Microsoft Graph a medida que exploras las carpetas, extrayendo columnas y metadatos personalizados vinculados a tus archivos." : "Navigate in real time through SharePoint document libraries (Drives). This explorer dynamically queries the Microsoft Graph API as you browse folders, extracting columns and custom metadata linked to your files."}
                        </p>
                      </div>
                    </div>
 
                    <DirectoryExplorer
                      siteId={resolvedSite.id}
                      isDemo={isDemo}
                      accessToken={accessToken}
                      language={language}
                    />
                  </motion.div>
                )}
 
                {/* TAB 2: RECURSIVE COMPLETE STRUCTURE MAPPER */}
                {activeTab === "mapper" && (
                  <motion.div
                    initial={{ opacity: 0, y: 15 }}
                    animate={{ opacity: 1, y: 0 }}
                    transition={{ duration: 0.3 }}
                    className="space-y-4"
                  >
                    {/* Control Panel or Loader */}
                    {mappingLoading ? (
                      <div className="bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded-lg p-12 flex flex-col items-center justify-center text-center space-y-4 shadow-sm transition-colors">
                        <RefreshCw className="w-10 h-10 text-blue-600 dark:text-blue-400 animate-spin" />
                        <div>
                          <h4 className="font-semibold text-slate-800 dark:text-slate-100 text-sm">{t.mappingLoadingText}</h4>
                          <p className="text-xs text-slate-500 dark:text-slate-400 mt-1 max-w-md mx-auto">
                            {language === "es" ? "Esto recorre recursivamente todas las bibliotecas de documentos, carpetas y archivos en Microsoft Graph, calculando la ruta completa para cada elemento y consolidando sus campos de metadatos." : "This recursively crawls all document libraries, folders, and files in Microsoft Graph, calculating the full path for each item and consolidating its metadata fields."}
                          </p>
                        </div>
                      </div>
                    ) : mappingError ? (
                      <div className="bg-rose-50 dark:bg-rose-950/20 border border-rose-200 dark:border-rose-900/30 rounded-lg p-6 text-center max-w-xl mx-auto space-y-4 shadow-sm transition-colors">
                        <AlertTriangle className="w-10 h-10 text-rose-600 dark:text-rose-400 mx-auto" />
                        <div>
                          <h4 className="font-semibold text-rose-900 dark:text-rose-300 text-sm">{t.errorMapHeader}</h4>
                          <p className="text-xs text-rose-700 dark:text-rose-400 mt-1">{mappingError}</p>
                        </div>
                        <button
                          onClick={handleMapCompleteStructure}
                          className="px-4 py-2 bg-rose-600 dark:bg-rose-800 text-white rounded-lg text-xs font-medium hover:bg-rose-700 dark:hover:bg-rose-700 cursor-pointer"
                        >
                          {t.btnRetryMapping}
                        </button>
                      </div>
                    ) : mappedItems.length === 0 ? (
                      <div className="bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded-lg p-12 text-center space-y-4 shadow-sm transition-colors">
                        <Database className="w-10 h-10 text-slate-400 dark:text-slate-500 mx-auto" />
                        <div>
                          <h4 className="font-semibold text-slate-800 dark:text-slate-100 text-sm">{t.noMapHeader}</h4>
                          <p className="text-xs text-slate-500 dark:text-slate-400 mt-1 max-w-sm mx-auto">
                            {t.noMapDesc}
                          </p>
                        </div>
                        <button
                          onClick={handleMapCompleteStructure}
                          className="px-4 py-2 bg-blue-600 dark:bg-blue-700 text-white rounded-lg text-xs font-semibold hover:bg-blue-700 dark:hover:bg-blue-600 cursor-pointer transition-all shadow-sm"
                        >
                          {t.btnGenerateMap}
                        </button>
                      </div>
                    ) : (
                      <div className="space-y-4">
                        {/* Brief alert showing success */}
                        {showMapSuccessAlert && (
                          <div className="bg-emerald-50/50 dark:bg-emerald-950/10 border border-emerald-200 dark:border-emerald-900/30 p-4 rounded-lg flex items-center justify-between gap-4 shadow-sm transition-colors">
                            <div className="flex items-center gap-2.5">
                              <CheckCircle className="w-5 h-5 text-emerald-600 dark:text-emerald-400" />
                              <span className="text-xs text-emerald-950 dark:text-emerald-300 font-medium">
                                {t.mapSuccessAlert} <strong className="text-emerald-900 dark:text-emerald-100 font-semibold">{mappedItems.length}</strong> {hasMoreThan15 ? t.trialLimitText : t.mapSuccessItems}.
                              </span>
                            </div>
                          </div>
                        )}
 
                        {/* Upselling Banner for Freemium Scan limit */}
                        {hasMoreThan15 && (
                          <div className="bg-gradient-to-r from-blue-600 via-blue-700 to-indigo-800 text-white p-5 rounded-lg shadow-md flex flex-col md:flex-row items-center justify-between gap-4 border border-blue-500">
                            <div className="flex items-center gap-3">
                              <div className="p-2 bg-white/10 rounded-lg">
                                <Sparkles className="w-6 h-6 text-yellow-300 animate-pulse shrink-0" />
                              </div>
                              <div>
                                <h4 className="font-bold text-sm text-yellow-100 flex items-center gap-1.5">
                                  {t.freePlanTitle}
                                </h4>
                                <p className="text-xs text-blue-100 mt-1">
                                  {t.freePlanDesc}
                                </p>
                              </div>
                            </div>
                            <button 
                              disabled 
                              className="px-4 py-2 bg-slate-300 dark:bg-slate-700 text-slate-500 dark:text-slate-400 font-bold text-xs rounded-md shadow shrink-0 cursor-not-allowed opacity-75"
                            >
                              {t.btnComingSoon}
                            </button>
                          </div>
                        )}

                        <StructureTable 
                          data={mappedItems} 
                          siteName={resolvedSite.displayName} 
                          language={language}
                        />
                      </div>
                    )}
                  </motion.div>
                )}

              </div>
            ) : (
              /* Landing Empty State with instructions */
              <div className="bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded-lg p-12 text-center max-w-xl mx-auto shadow-sm my-6 space-y-5 transition-colors">
                <div className="w-12 h-12 bg-blue-50 dark:bg-blue-950/40 text-blue-600 dark:text-blue-400 rounded flex items-center justify-center mx-auto">
                  <Database className="w-6 h-6" />
                </div>
                
                <div className="space-y-2">
                  <h3 className="font-semibold text-slate-800 dark:text-slate-100 text-lg">{language === "es" ? "Comienza validando el acceso a SharePoint" : "Start by validating SharePoint access"}</h3>
                  <p className="text-xs text-slate-500 dark:text-slate-400 leading-relaxed max-w-md mx-auto font-normal">
                    {isDemo 
                      ? (language === "es" ? "Haz clic en el botón 'Validar Acceso' arriba en el panel para simular la conexión a SharePoint en Modo Demo, o desactiva el Modo Demo para iniciar sesión con tus credenciales reales de Microsoft." : "Click on the 'Validate Access' button above in the panel to simulate the SharePoint connection in Demo Mode, or turn off Demo Mode to sign in with your real Microsoft credentials.")
                      : (language === "es" ? "Por favor, inicia sesión con tu cuenta de Microsoft (Paso A) y luego haz clic en 'Validar Acceso' (Paso B) para leer y mapear las bibliotecas del sitio." : "Please sign in with your Microsoft account (Step A) and then click 'Validate Access' (Step B) to read and map the site libraries.")
                    }
                  </p>
                </div>

                <div className="pt-2">
                  <button
                    id="init-demo-btn"
                    onClick={() => {
                      // Direct trigger to resolve in whatever current mode (defaults to Demo site path)
                      handleResolveSite();
                    }}
                    className="px-6 py-2.5 bg-blue-600 dark:bg-blue-700 hover:bg-blue-700 dark:hover:bg-blue-600 text-white text-xs font-semibold rounded-lg shadow-sm cursor-pointer transition-colors"
                  >
                    {language === "es" ? "Validar Acceso y Ver Sitio" : "Validate Access & View Site"}
                  </button>
                </div>
              </div>
            )}
          </div>
        ) : (
          /* CLASIFICADOR AI TAB VIEW */
          <div id="view-classifier" className="space-y-6">
            {!resolvedSite ? (
              /* If no site is resolved yet, request them to map structure */
              <div className="bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded-xl p-12 text-center max-w-xl mx-auto shadow-sm my-10 space-y-5 transition-colors">
                <div className="w-12 h-12 bg-purple-50 dark:bg-purple-950/40 text-purple-600 dark:text-purple-400 rounded-full flex items-center justify-center mx-auto shadow-inner">
                  <Brain className="w-6 h-6" />
                </div>
                <div className="space-y-2">
                  <h3 className="font-semibold text-slate-800 dark:text-slate-100 text-lg">{language === "es" ? "Conecta un sitio de SharePoint primero" : "Connect a SharePoint site first"}</h3>
                  <p className="text-xs text-slate-500 dark:text-slate-400 leading-relaxed max-w-md mx-auto font-normal">
                    {language === "es" ? "Para usar el Clasificador Inteligente con Gemini, primero necesitas validar tu acceso e introducir la URL de tu sitio de SharePoint en la pestaña Mapeador de Estructura." : "To use the Smart Classifier with Gemini, you first need to validate your access and enter your SharePoint site URL in the Structure Mapper tab."}
                  </p>
                </div>
                <div className="pt-2">
                  <button
                    onClick={() => setMainTab("mapper")}
                    className="px-6 py-2.5 bg-slate-900 dark:bg-slate-750 hover:bg-slate-800 dark:hover:bg-slate-650 text-white text-xs font-semibold rounded-lg shadow-sm cursor-pointer transition-all"
                  >
                    {language === "es" ? "Ir al Mapeador de Estructura" : "Go to Structure Mapper"}
                  </button>
                </div>
              </div>
            ) : mappedItems.length === 0 ? (
              /* If site is resolved but mappedItems is empty, request complete scan */
              <div className="bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded-xl p-12 text-center max-w-xl mx-auto shadow-sm my-10 space-y-5 transition-colors">
                <div className="w-12 h-12 bg-blue-50 dark:bg-blue-950/40 text-blue-600 dark:text-blue-400 rounded-full flex items-center justify-center mx-auto animate-pulse">
                  <Database className="w-6 h-6" />
                </div>
                <div className="space-y-2">
                  <h3 className="font-semibold text-slate-800 dark:text-slate-100 text-lg">{language === "es" ? "Escanear estructura para extraer archivos" : "Scan structure to extract files"}</h3>
                  <p className="text-xs text-slate-500 dark:text-slate-400 leading-relaxed max-w-md mx-auto font-normal font-sans">
                    {language === "es" ? <>Hemos verificado la conexión con el sitio <strong className="text-slate-800 dark:text-slate-100 font-semibold">{resolvedSite.displayName}</strong>, pero necesitamos realizar un escaneo para listar los archivos disponibles para clasificar.</> : <>We have verified the connection to the site <strong className="text-slate-800 dark:text-slate-100 font-semibold">{resolvedSite.displayName}</strong>, but we need to run a scan to list the files available for classification.</>}
                  </p>
                </div>
                <div className="pt-2">
                  <button
                    onClick={handleMapCompleteStructure}
                    disabled={mappingLoading}
                    className="px-6 py-2.5 bg-blue-600 dark:bg-blue-700 hover:bg-blue-700 dark:hover:bg-blue-600 disabled:bg-slate-200 dark:disabled:bg-slate-800 disabled:text-slate-400 dark:disabled:text-slate-500 text-white text-xs font-semibold rounded-lg shadow-sm cursor-pointer transition-all flex items-center gap-2 mx-auto"
                  >
                    {mappingLoading ? (
                      <>
                        <RefreshCw className="w-3.5 h-3.5 animate-spin" />
                        <span>{language === "es" ? "Escaneando carpetas de SharePoint..." : "Scanning SharePoint folders..."}</span>
                      </>
                    ) : (
                      <>
                        <Database className="w-3.5 h-3.5" />
                        <span>{language === "es" ? "Escanear y Cargar Archivos" : "Scan & Load Files"}</span>
                      </>
                    )}
                  </button>
                </div>
              </div>
            ) : (
              /* Full AI classifier interface with Split Layout */
              <div className="grid grid-cols-1 lg:grid-cols-12 gap-6 items-stretch">
                
                {/* Left Panel: Files and Folders List */}
                <div className="lg:col-span-4 bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded-xl p-5 shadow-sm flex flex-col space-y-4 max-h-[700px] transition-colors">
                  <div>
                    <h3 className="font-semibold text-slate-800 dark:text-slate-100 text-sm flex items-center gap-1.5">
                      <Layers className="w-4 h-4 text-slate-500 dark:text-slate-400" />
                      {language === "es" ? "Elementos Detectados" : "Detected Items"} ({mappedItems.length})
                    </h3>
                    <p className="text-[11px] text-slate-400 dark:text-slate-400 mt-0.5">{language === "es" ? "Selecciona un archivo o carpeta para clasificar mediante IA." : "Select a file or folder to classify using AI."}</p>
                  </div>

                  {/* Drag & Drop File Upload Area */}
                  <div 
                    onDragOver={(e) => { e.preventDefault(); setIsDragging(true); }}
                    onDragLeave={() => setIsDragging(false)}
                    onDrop={(e) => {
                      e.preventDefault();
                      setIsDragging(false);
                      handleLocalFileUpload(e.dataTransfer.files);
                    }}
                    className={`border-2 border-dashed rounded-xl p-3.5 text-center transition-all shrink-0 ${
                      isDragging 
                        ? "border-purple-500 bg-purple-50/40 dark:bg-purple-950/20 text-purple-700 dark:text-purple-400 font-medium" 
                        : "border-slate-200 dark:border-slate-750 hover:border-blue-400 bg-slate-50/20 dark:bg-slate-900/10 hover:bg-slate-50/70 dark:hover:bg-slate-900/30 text-slate-500 dark:text-slate-400"
                    }`}
                  >
                    <input 
                      type="file" 
                      id="classifier-file-upload" 
                      multiple 
                      className="hidden" 
                      onChange={(e) => handleLocalFileUpload(e.target.files)}
                    />
                    <label htmlFor="classifier-file-upload" className="cursor-pointer flex flex-col items-center justify-center gap-1.5">
                      <UploadCloud className={`w-5 h-5 ${isDragging ? 'text-purple-600 animate-bounce' : 'text-slate-400'}`} />
                      <div className="text-xs">
                        <span className="font-semibold text-blue-600 dark:text-blue-400 hover:text-blue-700 dark:hover:text-blue-300">{language === "es" ? "Cargar Archivos" : "Upload Files"}</span> {language === "es" ? "o arrastrar y soltar" : "or drag and drop"}
                      </div>
                      <p className="text-[9px] text-slate-400 dark:text-slate-500">{language === "es" ? "Modo Demo: se agregan localmente a la lista" : "Demo Mode: added locally to the list"}</p>
                    </label>
                  </div>

                  {/* Search bar & Export option inside list */}
                  <div className="flex flex-col gap-3 shrink-0">
                    <div className="relative">
                      <Search className="w-3.5 h-3.5 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2" />
                      <input
                        type="text"
                        placeholder={language === "es" ? "Filtrar por nombre..." : "Filter by name..."}
                        value={classifierSearch}
                        onChange={(e) => setClassifierSearch(e.target.value)}
                        className="w-full text-xs pl-8 pr-3 py-2 border border-slate-200 dark:border-slate-700 rounded-lg bg-slate-50/50 dark:bg-slate-900/20 focus:outline-none focus:ring-2 focus:ring-blue-500/10 focus:border-blue-500 dark:text-slate-100"
                      />
                    </div>
                    
                    {/* Elegant Export Action Menu */}
                    <div className="flex flex-col gap-2 bg-slate-50 dark:bg-slate-900/30 p-3 rounded-lg border border-slate-200 dark:border-slate-700">
                      <div className="flex items-center gap-1.5 text-[10px] font-bold uppercase tracking-wider text-slate-500 dark:text-slate-400">
                        <FileSpreadsheet className="w-4 h-4 text-emerald-600 shrink-0" />
                        <span>{language === "es" ? "Exportar Clasificación AI" : "Export AI Classification"}</span>
                      </div>
                      <div className="grid grid-cols-2 gap-2 mt-1">
                        <button
                          onClick={() => handleExport("csv")}
                          className="px-2.5 py-1.5 bg-white dark:bg-slate-850 border border-slate-200 dark:border-slate-700 hover:bg-slate-50 dark:hover:bg-slate-800 text-slate-700 dark:text-slate-300 text-[11px] font-semibold rounded-md shadow-xs cursor-pointer transition-colors flex items-center justify-center gap-1"
                          title={language === "es" ? "Descargar reporte consolidado en CSV" : "Download consolidated report in CSV"}
                        >
                          <span className="text-emerald-600 font-bold">CSV</span>
                          <span className="text-slate-400 dark:text-slate-500">{language === "es" ? "Descargar" : "Download"}</span>
                        </button>
                        <button
                          onClick={() => handleExport("json")}
                          className="px-2.5 py-1.5 bg-white dark:bg-slate-850 border border-slate-200 dark:border-slate-700 hover:bg-slate-50 dark:hover:bg-slate-800 text-slate-700 dark:text-slate-300 text-[11px] font-semibold rounded-md shadow-xs cursor-pointer transition-colors flex items-center justify-center gap-1"
                          title={language === "es" ? "Descargar volcado estructurado JSON" : "Download structured JSON dump"}
                        >
                          <span className="text-purple-600 font-bold">JSON</span>
                          <span className="text-slate-400 dark:text-slate-500">{language === "es" ? "Descargar" : "Download"}</span>
                        </button>
                      </div>
                    </div>
                  </div>

                  {/* Scrollable list of files and folders */}
                  <div className="flex-1 overflow-y-auto space-y-1.5 pr-1 divide-y divide-slate-50 dark:divide-slate-750">
                    {[...mappedItems]
                      .filter(item => item.name.toLowerCase().includes(classifierSearch.toLowerCase()))
                      .sort((a, b) => {
                        const dateA = a.lastModified ? new Date(a.lastModified).getTime() : 0;
                        const dateB = b.lastModified ? new Date(b.lastModified).getTime() : 0;
                        return dateB - dateA;
                      })
                      .map((file) => {
                        const isSelected = selectedFile?.id === file.id;
                        const hasMetaData = file.customMetadata && (file.customMetadata["AI Tags"] || file.customMetadata.MetaData);
                        const isFolder = file.type === "Carpeta";
                        return (
                          <button
                            key={file.id}
                            onClick={() => {
                              setSelectedFile(file);
                              setGeminiSuggestions([]);
                              setClassificationWarning(null);
                              setApplySuccess(null);
                              setApplyError(null);
                            }}
                            className={`w-full text-left p-3 rounded-lg text-xs transition-all flex items-start gap-3 cursor-pointer ${
                              isSelected
                                ? "bg-purple-50 dark:bg-purple-950/20 border border-purple-200 dark:border-purple-900/30 shadow-xs"
                                : "hover:bg-slate-50 dark:hover:bg-slate-750 bg-white dark:bg-slate-800"
                            }`}
                          >
                            {isFolder ? (
                              <Folder className={`w-4 h-4 mt-0.5 shrink-0 ${isSelected ? "text-purple-600" : "text-amber-500"}`} />
                            ) : (
                              <File className={`w-4 h-4 mt-0.5 shrink-0 ${isSelected ? "text-purple-600" : "text-blue-500"}`} />
                            )}
                            <div className="flex-1 min-w-0">
                              <p className={`font-semibold truncate ${isSelected ? "text-purple-950 dark:text-purple-200" : "text-slate-800 dark:text-slate-200"}`}>
                                {file.name}
                              </p>
                              <div className="flex items-center gap-1 text-[10px] text-slate-400 mt-0.5 font-normal">
                                <span className="truncate">{file.library}</span>
                                <span>•</span>
                                <span className="inline-flex items-center gap-0.5 px-1 py-0.2 rounded bg-slate-100 dark:bg-slate-700 text-[9px] text-slate-600 dark:text-slate-300 font-medium">
                                  {file.type === "Carpeta" ? (language === "es" ? "Carpeta" : "Folder") : (language === "es" ? "Archivo" : "File")}
                                </span>
                                <span>•</span>
                                <span>{new Date(file.lastModified).toLocaleDateString(language === "es" ? "es-ES" : "en-US", { month: 'short', day: 'numeric' })}</span>
                              </div>
                              {hasMetaData && (
                                <div className="mt-1.5 flex flex-wrap gap-0.5">
                                  {String(file.customMetadata["AI Tags"] || file.customMetadata.MetaData || "").split(/\s+/).filter(Boolean).map((tag, tIdx) => (
                                    <span key={tIdx} className="px-1.5 py-0.5 bg-slate-100 dark:bg-slate-900/40 border border-slate-200 dark:border-slate-700 text-slate-600 dark:text-slate-300 text-[9px] rounded font-mono font-medium scale-90 origin-left">
                                      {tag}
                                    </span>
                                  ))}
                                </div>
                              )}
                            </div>
                            <ChevronRight className="w-3.5 h-3.5 text-slate-300 shrink-0 self-center" />
                          </button>
                        );
                      })}

                    {mappedItems.filter(item => item.name.toLowerCase().includes(classifierSearch.toLowerCase())).length === 0 && (
                      <div className="p-8 text-center text-slate-400 dark:text-slate-500 text-xs">
                        {language === "es" ? "No se encontraron elementos con ese nombre." : "No items found with that name."}
                      </div>
                    )}
                  </div>
                </div>

                {/* Right Panel: Workspace */}
                <div className="lg:col-span-8 bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded-xl p-6 shadow-sm flex flex-col justify-between space-y-6 transition-colors">
                  {selectedFile === null ? (
                    /* Workspace Empty State */
                    <div className="flex-1 flex flex-col items-center justify-center text-center p-12 space-y-4">
                      <div className="w-16 h-16 bg-purple-50 dark:bg-purple-950/40 text-purple-600 dark:text-purple-400 rounded-full flex items-center justify-center shadow-inner">
                        <Sparkles className="w-7 h-7" />
                      </div>
                      <div className="max-w-sm">
                        <h4 className="font-semibold text-slate-800 dark:text-slate-100 text-base">{language === "es" ? "Asistente de Clasificación AI" : "AI Classification Assistant"}</h4>
                        <p className="text-xs text-slate-500 dark:text-slate-400 mt-1.5 leading-relaxed font-normal">
                          {language === "es" ? "Selecciona cualquier archivo del panel izquierdo para escanear sus metadatos e invocar a Gemini. El asistente te proporcionará los 5 hashtags más descriptivos listos para enviar a tu SharePoint." : "Select any file from the left panel to scan its metadata and invoke Gemini. The assistant will provide you with the 5 most descriptive hashtags ready to send to your SharePoint."}
                        </p>
                      </div>
                    </div>
                  ) : (
                    /* Active Workspace for selected file */
                    <div className="space-y-6 flex-1 flex flex-col justify-between">
                      <div className="space-y-5">
                        {/* File summary bar */}
                        <div className="p-4 bg-slate-50 dark:bg-slate-900/30 border border-slate-200 dark:border-slate-700 rounded-xl flex items-center justify-between gap-4 transition-colors">
                          <div className="flex items-start gap-3">
                            <div className={`p-2.5 bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded-lg shadow-xs mt-0.5 shrink-0 ${selectedFile.type === "Carpeta" ? "text-amber-500" : "text-blue-500"}`}>
                              {selectedFile.type === "Carpeta" ? (
                                <Folder className="w-5 h-5" />
                              ) : (
                                <FileText className="w-5 h-5" />
                              )}
                            </div>
                            <div className="text-xs">
                              <h4 className="font-bold text-slate-800 dark:text-slate-100 break-all leading-snug">{selectedFile.name}</h4>
                              <p className="text-slate-400 mt-0.5 font-normal">
                                {language === "es" ? "Biblioteca:" : "Library:"} <strong className="text-slate-600 dark:text-slate-350 font-semibold">{selectedFile.library}</strong> | 
                                {language === "es" ? "Ruta:" : "Path:"} <code className="text-[10px] bg-slate-100 dark:bg-slate-900 px-1 py-0.5 rounded font-mono text-slate-600 dark:text-slate-400">{selectedFile["Estructura actual SharePoint"]}</code>
                              </p>
                            </div>
                          </div>
                        </div>

                        {/* Current columns and tags */}
                        <div className="space-y-1.5">
                          <span className="text-[11px] font-bold text-slate-400 dark:text-slate-450 uppercase tracking-wider block">{language === "es" ? "Estado de Metadatos Actual" : "Current Metadata Status"}</span>
                          <div className="p-4 bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded-xl text-xs space-y-2 transition-colors">
                            <div className="flex items-center justify-between">
                              <span className="font-medium text-slate-500 dark:text-slate-400">{language === "es" ? "Columna SharePoint" : "SharePoint Column"} <code className="text-blue-600 dark:text-blue-400 bg-blue-50 dark:bg-blue-950/30 font-mono text-[10px] px-1 rounded font-bold">AI Tags</code>:</span>
                              {(selectedFile.customMetadata["AI Tags"] || selectedFile.customMetadata.MetaData) ? (
                                <span className="text-emerald-700 dark:text-emerald-400 font-bold bg-emerald-50 dark:bg-emerald-950/20 px-2 py-0.5 border border-emerald-100 dark:border-emerald-900/30 rounded text-[10px]">
                                  {language === "es" ? "Clasificado" : "Classified"}
                                </span>
                              ) : (
                                <span className="text-amber-700 dark:text-amber-400 font-bold bg-amber-50 dark:bg-amber-950/20 px-2 py-0.5 border border-amber-100 dark:border-amber-900/30 rounded text-[10px]">
                                  {language === "es" ? "Sin clasificar" : "Unclassified"}
                                </span>
                              )}
                            </div>
                            <div className="font-mono text-xs bg-slate-50 dark:bg-slate-900/40 p-2.5 rounded border border-slate-200 dark:border-slate-700 text-slate-800 dark:text-slate-200 break-all leading-relaxed min-h-[36px] transition-colors">
                              {(selectedFile.customMetadata["AI Tags"] || selectedFile.customMetadata.MetaData) || (language === "es" ? "— (No contiene etiquetas actualmente)" : "— (Currently contains no tags)")}
                            </div>
                          </div>
                        </div>

                        {/* Invoking Gemini and suggestions */}
                        <div className="space-y-4">
                          <div className="flex items-center justify-between gap-4 flex-wrap">
                            <div className="flex items-center gap-2">
                              <span className="text-[11px] font-bold text-slate-400 dark:text-slate-450 uppercase tracking-wider">{language === "es" ? "Metadatos de Clasificación AI (Gemini)" : "AI Classification Metadata (Gemini)"}</span>
                              <span className={`text-[9px] font-mono font-bold px-2 py-0.5 rounded-full transition-colors ${aiCallCount >= 90 ? 'bg-rose-100 dark:bg-rose-950/40 text-rose-800 dark:text-rose-400 animate-pulse' : 'bg-slate-100 dark:bg-slate-700 text-slate-600 dark:text-slate-300'}`}>
                                {aiCallCount}/100 {language === "es" ? "llamadas" : "calls"}
                              </span>
                            </div>
                            <button
                              id="suggest-categories-btn"
                              onClick={() => handleSuggestCategories(selectedFile)}
                              disabled={isClassifying || aiCallCount >= 100}
                              className="inline-flex items-center gap-1.5 px-3.5 py-2 bg-purple-600 hover:bg-purple-700 dark:bg-purple-700 dark:hover:bg-purple-600 disabled:bg-slate-200 dark:disabled:bg-slate-800 disabled:text-slate-400 dark:disabled:text-slate-500 text-white text-xs font-semibold rounded-lg transition-all shadow-sm cursor-pointer"
                            >
                              {isClassifying ? (
                                <>
                                  <RefreshCw className="w-3.5 h-3.5 animate-spin" />
                                  <span>{language === "es" ? "Clasificando con Gemini..." : "Classifying with Gemini..."}</span>
                                </>
                              ) : (
                                <>
                                  <Brain className="w-3.5 h-3.5" />
                                  <span>{language === "es" ? "Sugerir Categorías" : "Suggest Categories"}</span>
                                </>
                              )}
                            </button>
                          </div>

                          {/* 100 Calls Max Block Alert */}
                          {aiCallCount >= 100 && (
                            <div className="p-3.5 bg-rose-50 dark:bg-rose-950/20 border border-rose-200 dark:border-rose-900/30 rounded-xl text-xs text-rose-800 dark:text-rose-400 flex gap-2.5 shadow-xs transition-colors">
                              <AlertCircle className="w-4 h-4 shrink-0 mt-0.5 text-rose-600 animate-bounce" />
                              <div className="space-y-1">
                                <p className="font-bold">{language === "es" ? "Límite de procesamiento de IA alcanzado por seguridad del servidor." : "AI processing limit reached for server safety."}</p>
                                <p className="text-rose-700 dark:text-rose-450 font-normal leading-normal">{language === "es" ? "Si necesitas procesar volúmenes masivos, contacta a RIEA para una solución dedicada corporativa." : "If you need to process massive volumes, contact RIEA for a dedicated enterprise solution."}</p>
                              </div>
                            </div>
                          )}

                          {/* 90% Quota Subtle Alert */}
                          {aiCallCount >= 90 && aiCallCount < 100 && (
                            <div className="p-3 bg-amber-50 dark:bg-amber-950/20 border border-amber-200 dark:border-amber-900/30 rounded-xl text-xs text-amber-800 dark:text-amber-400 flex gap-2.5 shadow-xs transition-colors">
                              <Sparkles className="w-4 h-4 shrink-0 mt-0.5 text-amber-600 animate-pulse" />
                              <div>
                                <p className="font-semibold">{language === "es" ? `Has consumido el ${aiCallCount}% de tu cuota de procesamiento optimizada para esta sesión.` : `You have consumed ${aiCallCount}% of your optimized processing quota for this session.`}</p>
                              </div>
                            </div>
                          )}

                          {/* Warning / Backups notification for Gemini API keys if any */}
                          {classificationWarning && (
                            <div className="p-3 bg-amber-50 dark:bg-amber-950/20 border border-amber-200 dark:border-amber-900/30 rounded-lg text-xs text-amber-800 dark:text-amber-400 flex gap-2 transition-colors">
                              <AlertTriangle className="w-4 h-4 shrink-0 mt-0.5 text-amber-600" />
                              <span>{classificationWarning}</span>
                            </div>
                          )}

                          {/* Editable suggested inputs */}
                          <div className="space-y-3 p-4 bg-purple-50/20 dark:bg-purple-950/5 border border-purple-100 dark:border-purple-900/20 rounded-xl transition-colors">
                            <div className="text-xs text-purple-950 dark:text-purple-200 font-semibold flex items-center justify-between">
                              <span>{language === "es" ? "Revisa, edita o complementa los 5 Hashtags de destino:" : "Review, edit or supplement the 5 target Hashtags:"}</span>
                              <span className="text-[10px] font-mono font-normal text-slate-400 dark:text-slate-500">{language === "es" ? "Total: 5 campos asignados" : "Total: 5 fields assigned"}</span>
                            </div>
                            
                            <div className="grid grid-cols-1 md:grid-cols-5 gap-2">
                              {editableTags.map((tag, idx) => (
                                <div key={idx} className="relative flex items-center">
                                  <span className="absolute left-2.5 text-slate-400 dark:text-slate-500 font-mono text-xs font-semibold select-none">#</span>
                                  <input
                                    type="text"
                                    value={tag.startsWith("#") ? tag.substring(1) : tag}
                                    onChange={(e) => {
                                      const newVal = e.target.value.replace(/\s+/g, ""); // No spaces in hashtags
                                      const updated = [...editableTags];
                                      updated[idx] = newVal ? "#" + newVal : "";
                                      setEditableTags(updated);
                                    }}
                                    placeholder={language === "es" ? `Etiqueta ${idx + 1}` : `Tag ${idx + 1}`}
                                    className="w-full text-xs font-mono pl-6 pr-1.5 py-2 border border-slate-200 dark:border-slate-700 bg-white dark:bg-slate-800 rounded-lg focus:outline-none focus:ring-2 focus:ring-purple-500/15 focus:border-purple-500 font-medium text-slate-800 dark:text-slate-100 shadow-inner"
                                  />
                                </div>
                              ))}
                            </div>
                            <p className="text-[10px] text-slate-400 dark:text-slate-450 leading-normal">
                              {language === "es" ? "* Al guardar, los hashtags completados se unificarán con espacios (ej: #Contrato #Finanzas #Proveedor) y se guardarán en la columna de destino de SharePoint." : "* When saving, the completed hashtags will be unified with spaces (e.g., #Contract #Finance #Vendor) and saved in the target SharePoint column."}
                            </p>
                          </div>
                        </div>
                      </div>

                      {/* Apply button and banners */}
                      <div className="space-y-4 pt-4 border-t border-slate-100 dark:border-slate-700 shrink-0 transition-colors">
                        {/* Status Banners */}
                        <AnimatePresence mode="wait">
                          {applySuccess && (
                            <motion.div
                              key="apply-success"
                              initial={{ opacity: 0, y: 5 }}
                              animate={{ opacity: 1, y: 0 }}
                              exit={{ opacity: 0, y: 5 }}
                              className="p-3 bg-emerald-50 dark:bg-emerald-950/20 border border-emerald-200 dark:border-emerald-900/30 text-emerald-800 dark:text-emerald-400 text-xs rounded-lg flex items-center gap-2 font-medium transition-colors"
                            >
                              <CheckCircle className="w-4 h-4 shrink-0 text-emerald-600" />
                              <span>{applySuccess}</span>
                            </motion.div>
                          )}

                          {applyError && (
                            <motion.div
                              key="apply-error"
                              initial={{ opacity: 0, y: 5 }}
                              animate={{ opacity: 1, y: 0 }}
                              exit={{ opacity: 0, y: 5 }}
                              className="p-3 bg-rose-50 dark:bg-rose-950/20 border border-rose-200 dark:border-rose-900/30 text-rose-800 dark:text-rose-400 text-xs rounded-lg flex items-center gap-2 font-medium transition-colors"
                            >
                              <AlertCircle className="w-4 h-4 shrink-0 text-rose-600" />
                              <span>{applyError}</span>
                            </motion.div>
                          )}
                        </AnimatePresence>

                        <div className="flex flex-col md:flex-row items-stretch md:items-center justify-between gap-4">
                          {/* Admin Note on left */}
                          <div className="text-[10px] text-slate-400 dark:text-slate-400 max-w-md leading-normal font-normal">
                            {language === "es" ? (
                              <><strong>Nota del Administrador:</strong> Esta operación requiere que el sitio de SharePoint contenga una columna personalizada de tipo <em>Múltiples líneas de texto (Multiple lines of text)</em> con formato de <em>Texto sin formato (Plain text)</em> nombrada exactamente <code className="font-bold text-slate-700 dark:text-slate-350 bg-slate-100 dark:bg-slate-900 px-1 py-0.5 rounded">AI Tags</code>.</>
                            ) : (
                              <><strong>Administrator Note:</strong> This operation requires that the SharePoint site contains a custom column of type <em>Multiple lines of text</em> with <em>Plain text</em> formatting named exactly <code className="font-bold text-slate-700 dark:text-slate-350 bg-slate-100 dark:bg-slate-900 px-1 py-0.5 rounded">AI Tags</code>.</>
                            )}
                          </div>

                          {/* Save trigger button */}
                          <button
                            id="apply-metadata-btn"
                            disabled={isApplyingMetadata}
                            onClick={() => {
                              const finalString = editableTags
                                .map(t => t.trim())
                                .filter(Boolean)
                                .map(t => t.startsWith("#") ? t : "#" + t)
                                .join(" ");
                              handleApplyMetadata(finalString);
                            }}
                            className="inline-flex items-center justify-center gap-1.5 px-6 py-2.5 bg-slate-900 dark:bg-slate-700 hover:bg-slate-800 dark:hover:bg-slate-600 disabled:bg-slate-200 dark:disabled:bg-slate-800 text-white disabled:text-slate-400 dark:disabled:text-slate-500 text-xs font-semibold rounded-lg shadow-sm transition-all cursor-pointer whitespace-nowrap shrink-0"
                          >
                            {isApplyingMetadata ? (
                              <>
                                <RefreshCw className="w-3.5 h-3.5 animate-spin" />
                                <span>{language === "es" ? "Guardando en SharePoint..." : "Saving to SharePoint..."}</span>
                              </>
                            ) : (
                              <>
                                <Check className="w-3.5 h-3.5 text-emerald-400" />
                                <span>{language === "es" ? "Aplicar a SharePoint" : "Apply to SharePoint"}</span>
                              </>
                            )}
                          </button>
                        </div>
                      </div>
                    </div>
                  )}
                </div>

              </div>
            )}
          </div>
        )}

      </main>

      {/* FOOTER BAR */}
      <footer id="app-footer" className="bg-white dark:bg-slate-800 border-t border-slate-200 dark:border-slate-700 py-5 px-6 text-center text-xs text-slate-500 dark:text-slate-400 shrink-0 font-medium transition-colors">
        <div className="max-w-7xl mx-auto flex flex-col sm:flex-row items-center justify-between gap-4">
          <span>Microsoft Graph OAuth2 & SharePoint Structure Mapper - Entra ID Integration</span>
          <div className="flex gap-4">
            <a href="https://learn.microsoft.com/en-us/graph/api/resources/sharepoint" target="_blank" rel="noopener noreferrer" className="hover:text-blue-600 dark:hover:text-blue-400 hover:underline flex items-center gap-1">
              {language === "es" ? "Documentación de Graph API" : "Graph API Documentation"}
              <ExternalLink className="w-3.5 h-3.5" />
            </a>
          </div>
        </div>
      </footer>
    </div>
  );
}
