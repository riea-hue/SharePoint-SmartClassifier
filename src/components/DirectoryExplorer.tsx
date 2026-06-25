import React, { useState, useEffect } from "react";
import { Folder, File, ChevronRight, HardDrive, RefreshCw, Layers, Calendar, HardDriveUpload, CheckSquare, Search } from "lucide-react";

interface Drive {
  id: string;
  name: string;
  description: string;
}

interface DriveItem {
  id: string;
  name: string;
  folder?: any;
  file?: any;
  size: number;
  webUrl: string;
  lastModifiedDateTime: string;
  metadata?: Record<string, string>;
}

interface DirectoryExplorerProps {
  siteId: string;
  isDemo: boolean;
  accessToken: string | null;
  onSelectFile?: (item: DriveItem, path: string) => void;
  language: Language;
}

import { Language, translations } from "../translations";

export default function DirectoryExplorer({ siteId, isDemo, accessToken, onSelectFile, language }: DirectoryExplorerProps) {
  const [drives, setDrives] = useState<Drive[]>([]);
  const [selectedDrive, setSelectedDrive] = useState<Drive | null>(null);
  const [items, setItems] = useState<DriveItem[]>([]);
  const [loadingDrives, setLoadingDrives] = useState(false);
  const [loadingItems, setLoadingItems] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const t = translations[language];
  
  // Breadcrumb / folder stack representation: array of { id, name }
  const [folderStack, setFolderStack] = useState<{ id: string; name: string }[]>([]);
  const [searchQuery, setSearchQuery] = useState("");

  // Load drives on mount or site changes
  useEffect(() => {
    if (!siteId) return;
    
    const fetchDrives = async () => {
      setLoadingDrives(true);
      setError(null);
      setSelectedDrive(null);
      setItems([]);
      setFolderStack([]);
      
      try {
        const headers: Record<string, string> = {};
        if (accessToken) {
          headers["Authorization"] = `Bearer ${accessToken}`;
        }
        
        const response = await fetch(`/api/sharepoint/sites/${siteId}/drives?isDemo=${isDemo}`, {
          headers
        });
        
        if (!response.ok) {
          const errData = await response.json();
          throw new Error(errData.message || "No se pudieron cargar las bibliotecas de documentos.");
        }
        
        const data = await response.json();
        setDrives(data.drives || []);
        if (data.drives && data.drives.length > 0) {
          handleSelectDrive(data.drives[0]);
        }
      } catch (err: any) {
        console.error(err);
        setError(err.message || "Ocurrió un error al cargar las bibliotecas.");
      } finally {
        setLoadingDrives(false);
      }
    };

    fetchDrives();
  }, [siteId, isDemo, accessToken]);

  // Load items of current folder
  const fetchItems = async (driveId: string, folderId: string | null = null) => {
    setLoadingItems(true);
    setError(null);
    try {
      const headers: Record<string, string> = {};
      if (accessToken) {
        headers["Authorization"] = `Bearer ${accessToken}`;
      }
      
      let url = `/api/sharepoint/drives/${driveId}/items?isDemo=${isDemo}`;
      if (folderId) {
        url += `&folderId=${folderId}`;
      }
      
      const response = await fetch(url, { headers });
      if (!response.ok) {
        const errData = await response.json();
        throw new Error(errData.message || "Error al cargar los archivos.");
      }
      
      const data = await response.json();
      setItems(data.items || []);
    } catch (err: any) {
      console.error(err);
      setError(err.message || "Error al conectar con SharePoint.");
    } finally {
      setLoadingItems(false);
    }
  };

  const handleSelectDrive = (drive: Drive) => {
    setSelectedDrive(drive);
    setFolderStack([]);
    fetchItems(drive.id);
  };

  const handleFolderClick = (folder: DriveItem) => {
    const newStack = [...folderStack, { id: folder.id, name: folder.name }];
    setFolderStack(newStack);
    if (selectedDrive) {
      fetchItems(selectedDrive.id, folder.id);
    }
  };

  const handleBreadcrumbClick = (index: number) => {
    if (index === -1) {
      // Go to drive root
      setFolderStack([]);
      if (selectedDrive) {
        fetchItems(selectedDrive.id);
      }
    } else {
      // Slice the stack to go back to target folder
      const newStack = folderStack.slice(0, index + 1);
      setFolderStack(newStack);
      if (selectedDrive) {
        fetchItems(selectedDrive.id, newStack[newStack.length - 1].id);
      }
    }
  };

  const formatSize = (bytes: number) => {
    if (bytes === 0) return "—";
    const k = 1024;
    const sizes = ["Bytes", "KB", "MB", "GB"];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(1)) + " " + sizes[i];
  };

  const formatDate = (isoString: string) => {
    return new Date(isoString).toLocaleString(language === "es" ? "es-ES" : "en-US", {
      year: "numeric",
      month: "short",
      day: "numeric",
      hour: "2-digit",
      minute: "2-digit"
    });
  };

  const getFullItemPath = (item: DriveItem) => {
    const breadcrumbPath = folderStack.map(f => f.name).join("/");
    const driveName = selectedDrive ? selectedDrive.name : "Documentos";
    return breadcrumbPath 
      ? `${driveName}/${breadcrumbPath}/${item.name}`
      : `${driveName}/${item.name}`;
  };

  const filteredItems = items
    .filter(item => item.name.toLowerCase().includes(searchQuery.toLowerCase()))
    .sort((a, b) => {
      const isFolderA = !!a.folder;
      const isFolderB = !!b.folder;
      
      if (isFolderA && !isFolderB) return -1;
      if (!isFolderA && isFolderB) return 1;
      
      return a.name.localeCompare(b.name);
    });

  return (
    <div id="directory-explorer-wrapper" className="bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded-lg overflow-hidden shadow-sm flex flex-col h-[520px] transition-colors">
      {/* Selector de Bibliotecas */}
      <div className="bg-slate-50 dark:bg-slate-900/60 border-b border-slate-200 dark:border-slate-700 px-5 py-4 flex flex-col md:flex-row gap-4 items-stretch md:items-center justify-between shrink-0">
        <div className="flex items-center gap-2.5">
          <HardDrive className="w-5 h-5 text-blue-600 dark:text-blue-400" />
          <span className="font-semibold text-slate-800 dark:text-slate-100">
            {language === "es" ? "Bibliotecas de Documentos" : "Document Libraries"}
          </span>
          {loadingDrives && <RefreshCw className="w-4 h-4 text-blue-600 dark:text-blue-400 animate-spin" />}
        </div>
        
        <div className="flex gap-2 overflow-x-auto max-w-full pb-1 md:pb-0 scrollbar-thin">
          {drives.map(drive => (
            <button
              id={`drive-tab-${drive.id}`}
              key={drive.id}
              onClick={() => handleSelectDrive(drive)}
              className={`px-3 py-1.5 rounded-md text-xs font-medium border transition-all whitespace-nowrap cursor-pointer ${
                selectedDrive?.id === drive.id
                  ? "bg-blue-600 text-white border-blue-600 shadow-xs"
                  : "bg-white dark:bg-slate-800 text-slate-600 dark:text-slate-300 border-slate-200 dark:border-slate-700 hover:bg-slate-50 dark:hover:bg-slate-750"
              }`}
            >
              {drive.name}
            </button>
          ))}
          {drives.length === 0 && !loadingDrives && (
            <span className="text-xs text-slate-400 dark:text-slate-500 self-center">
              {language === "es" ? "No se encontraron bibliotecas" : "No libraries found"}
            </span>
          )}
        </div>
      </div>

      {/* Barra de navegación de carpetas (Breadcrumbs) y búsqueda */}
      <div className="px-5 py-3 bg-white dark:bg-slate-800 border-b border-slate-100 dark:border-slate-700 flex flex-col md:flex-row gap-3 items-stretch md:items-center justify-between shrink-0">
        {/* Breadcrumbs */}
        <div className="flex items-center gap-1.5 flex-wrap text-xs text-slate-500 dark:text-slate-400 overflow-hidden font-medium">
          <button
            onClick={() => handleBreadcrumbClick(-1)}
            className="text-blue-600 dark:text-blue-400 hover:underline hover:text-blue-700 dark:hover:text-blue-300 font-semibold cursor-pointer"
          >
            {selectedDrive ? selectedDrive.name : (language === "es" ? "Raíz" : "Root")}
          </button>
          
          {folderStack.map((folder, index) => (
            <React.Fragment key={folder.id}>
              <ChevronRight className="w-3.5 h-3.5 text-slate-400 dark:text-slate-500 shrink-0" />
              <button
                onClick={() => handleBreadcrumbClick(index)}
                className={`hover:underline shrink-0 max-w-[140px] truncate cursor-pointer ${
                  index === folderStack.length - 1 ? "text-slate-800 dark:text-slate-100 font-semibold" : "text-blue-600 dark:text-blue-400"
                }`}
              >
                {folder.name}
              </button>
            </React.Fragment>
          ))}
        </div>

        {/* Búsqueda */}
        <div className="relative max-w-full md:max-w-[240px]">
          <Search className="w-4 h-4 text-slate-400 dark:text-slate-500 absolute left-3 top-1/2 -translate-y-1/2" />
          <input
            type="text"
            placeholder={t.explorerSearchPlaceholder || "Filtrar archivos..."}
            value={searchQuery}
            onChange={(e) => setSearchQuery(e.target.value)}
            className="w-full text-xs pl-9 pr-4 py-1.5 border border-slate-200 dark:border-slate-700 rounded-md bg-white dark:bg-slate-800 text-slate-800 dark:text-slate-100 focus:outline-none focus:ring-2 focus:ring-blue-500/10 focus:border-blue-500"
          />
        </div>
      </div>

      {/* Lista de archivos / folders */}
      <div className="flex-1 overflow-y-auto bg-slate-50/50 dark:bg-slate-900/20 p-3">
        {loadingItems ? (
          <div className="flex flex-col items-center justify-center h-full text-slate-400 dark:text-slate-500 gap-2.5">
            <RefreshCw className="w-8 h-8 text-blue-600 dark:text-blue-400 animate-spin" />
            <span className="text-xs font-medium">
              {language === "es" ? "Conectando y cargando archivos de SharePoint..." : "Connecting and loading SharePoint files..."}
            </span>
          </div>
        ) : error ? (
          <div className="flex flex-col items-center justify-center h-full text-center px-6 py-12 bg-rose-50 dark:bg-rose-950/20 rounded-lg border border-rose-100 dark:border-rose-900/30 max-w-lg mx-auto my-6 shadow-sm">
            <div className="w-12 h-12 rounded-full bg-rose-100 dark:bg-rose-900/40 flex items-center justify-center text-rose-600 dark:text-rose-400 mb-3.5">
              <Layers className="w-6 h-6" />
            </div>
            <h4 className="text-sm font-semibold text-rose-800 dark:text-rose-200">
              {language === "es" ? "No se pudieron cargar los datos" : "Failed to load data"}
            </h4>
            <p className="text-xs text-rose-600 dark:text-rose-400 mt-1.5">{error}</p>
          </div>
        ) : filteredItems.length === 0 ? (
          <div className="flex flex-col items-center justify-center h-full text-slate-400 dark:text-slate-500 py-16">
            <Folder className="w-10 h-10 text-slate-300 dark:text-slate-600 mb-2" />
            <p className="text-xs font-medium">{t.emptyDirectory || "Este directorio está vacío."}</p>
          </div>
        ) : (
          <div className="bg-white dark:bg-slate-800 border border-slate-200/60 dark:border-slate-700 rounded-lg overflow-x-auto overflow-y-auto max-h-[500px] shadow-xs">
            <table className="w-full text-left border-collapse min-w-[800px]">
              <thead>
                <tr className="bg-slate-50/70 dark:bg-slate-900/80 border-b border-slate-200 dark:border-slate-700 text-xs text-slate-500 dark:text-slate-400 font-semibold uppercase tracking-wider sticky top-0 z-10 font-bold">
                  <th className="px-4 py-3 bg-slate-50/90 dark:bg-slate-900/90 w-[35%]">{language === "es" ? "Nombre" : "Name"}</th>
                  <th className="px-4 py-3 bg-slate-50/90 dark:bg-slate-900/90 w-[20%]">{language === "es" ? "Modificado" : "Modified"}</th>
                  <th className="px-4 py-3 bg-slate-50/90 dark:bg-slate-900/90 w-[15%]">{language === "es" ? "Tamaño" : "Size"}</th>
                  <th className="px-4 py-3 bg-slate-50/90 dark:bg-slate-900/90 w-[30%]">{language === "es" ? "Columnas Personalizadas (Metadatos)" : "Custom Columns (Metadata)"}</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100 dark:divide-slate-700 text-xs text-slate-600 dark:text-slate-300">
                {filteredItems.map(item => {
                  const isFolder = !!item.folder;
                  const itemPath = getFullItemPath(item);

                  return (
                    <tr
                      key={item.id}
                      className="hover:bg-slate-50/65 dark:hover:bg-slate-700/20 transition-colors group cursor-pointer"
                      onClick={() => isFolder ? handleFolderClick(item) : onSelectFile?.(item, itemPath)}
                    >
                      {/* Nombre */}
                      <td className="px-4 py-3.5 font-medium text-slate-800 dark:text-slate-100">
                        <div className="flex items-center gap-2.5 max-w-[300px] truncate">
                          {isFolder ? (
                            <Folder className="w-4 h-4 text-amber-500 shrink-0 group-hover:scale-105 transition-transform" />
                          ) : (
                            <File className="w-4 h-4 text-blue-500 shrink-0" />
                          )}
                          <span className="truncate hover:text-blue-600 dark:hover:text-blue-400 group-hover:underline" title={item.name}>
                            {item.name}
                          </span>
                        </div>
                      </td>

                      {/* Fecha de Modificacion */}
                      <td className="px-4 py-3.5 text-slate-400 dark:text-slate-500 whitespace-nowrap">
                        <div className="flex items-center gap-1.5">
                          <Calendar className="w-3.5 h-3.5 text-slate-300 dark:text-slate-600" />
                          {formatDate(item.lastModifiedDateTime)}
                        </div>
                      </td>

                      {/* Tamaño */}
                      <td className="px-4 py-3.5 text-slate-500 dark:text-slate-400 whitespace-nowrap font-mono text-[11px]">
                        {formatSize(item.size)}
                      </td>

                      {/* Metadatos Personalizados */}
                      <td className="px-4 py-3.5">
                        <div className="max-h-[64px] overflow-y-auto pr-1 flex flex-col gap-1 scrollbar-thin max-w-[240px]">
                          {item.metadata ? (
                            Object.entries(item.metadata).map(([key, value]) => (
                              <div
                                key={key}
                                className="px-2 py-0.5 bg-slate-50 dark:bg-slate-900 border border-slate-200 dark:border-slate-700 text-slate-600 dark:text-slate-400 text-[10px] font-mono rounded flex items-center justify-between gap-1"
                                title={`${key}: ${value}`}
                              >
                                <span className="font-semibold text-slate-700 dark:text-slate-400 shrink-0">{key}:</span>
                                <span className="text-slate-900 dark:text-slate-200 truncate">{value}</span>
                              </div>
                            ))
                          ) : isFolder ? (
                            <span className="text-[10px] text-slate-400 bg-slate-100 dark:bg-slate-900 px-1.5 py-0.5 rounded font-medium italic self-start">
                              {language === "es" ? "Directorio" : "Directory"}
                            </span>
                          ) : (
                            <span className="text-[10px] text-slate-300 dark:text-slate-500 italic font-medium">{t.noMetadataText || "Sin columnas"}</span>
                          )}
                        </div>
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        )}
      </div>

      {/* Indicador de ayuda */}
      <div className="bg-slate-50 dark:bg-slate-900/60 px-5 py-3 border-t border-slate-100 dark:border-slate-700 flex items-center justify-between text-xs text-slate-500 dark:text-slate-400 shrink-0 font-medium transition-colors">
        <span>
          {language === "es" ? "* Haz clic en una carpeta para explorarla." : "* Click on a folder to explore it."}
        </span>
        <span>
          {language === "es" ? `Mostrando ${filteredItems.length} elementos` : `Showing ${filteredItems.length} items`}
        </span>
      </div>
    </div>
  );
}

