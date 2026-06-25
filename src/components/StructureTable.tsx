import React, { useState } from "react";
import { Download, Search, FileJson, Copy, Check, Filter, Layers, HelpCircle, File, Folder } from "lucide-react";
import { Language, translations } from "../translations";

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

interface StructureTableProps {
  data: MappedItem[];
  siteName: string;
  language: Language;
}

export default function StructureTable({ data, siteName, language }: StructureTableProps) {
  const [searchTerm, setSearchTerm] = useState("");
  const [typeFilter, setTypeFilter] = useState<"Todos" | "Archivo" | "Carpeta">("Todos");
  const [libraryFilter, setLibraryFilter] = useState<string>("Todas");
  const [copiedId, setCopiedId] = useState<string | null>(null);
  const t = translations[language];

  // Extract unique libraries for filter dropdown
  const uniqueLibraries = ["Todas", ...Array.from(new Set(data.map((item) => item.library)))];

  // Filtering logic
  const filteredData = data.filter((item) => {
    const matchesSearch =
      item.name.toLowerCase().includes(searchTerm.toLowerCase()) ||
      item["Estructura actual SharePoint"].toLowerCase().includes(searchTerm.toLowerCase());

    const matchesType = typeFilter === "Todos" || item.type === typeFilter;
    const matchesLibrary = libraryFilter === "Todas" || item.library === libraryFilter;

    return matchesSearch && matchesType && matchesLibrary;
  });

  const handleCopyPath = (path: string, id: string) => {
    navigator.clipboard.writeText(path);
    setCopiedId(id);
    setTimeout(() => setCopiedId(null), 2000);
  };

  const handleDownloadJson = () => {
    const fileData = JSON.stringify(
      {
        site: siteName,
        exportDate: new Date().toISOString(),
        totalElements: filteredData.length,
        structure: filteredData,
      },
      null,
      2
    );

    const blob = new Blob([fileData], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    
    // Clean site name for filename
    const safeName = siteName.toLowerCase().replace(/[^a-z0-9]/gi, "_");
    link.download = `sharepoint_structure_${safeName || "mapped"}.json`;
    
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
  };

  const formatSize = (bytes: number) => {
    if (bytes === 0) return "—";
    const k = 1024;
    const sizes = ["Bytes", "KB", "MB", "GB"];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(1)) + " " + sizes[i];
  };

  const formatDate = (isoString: string) => {
    return new Date(isoString).toLocaleDateString(language === "es" ? "es-ES" : "en-US", {
      year: "numeric",
      month: "short",
      day: "numeric",
    });
  };

  return (
    <div id="structure-table-container" className="bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded-lg overflow-hidden shadow-sm flex flex-col space-y-4 p-5 transition-colors">
      {/* Header section */}
      <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
        <div>
          <h3 className="font-semibold text-slate-800 dark:text-slate-100 text-lg flex items-center gap-2">
            <Layers className="w-5 h-5 text-blue-600 dark:text-blue-400" />
            {t.tableTitle}
          </h3>
          <p className="text-xs text-slate-500 dark:text-slate-400 mt-1">
            {t.tableSubtitle}
          </p>
        </div>

        <div className="flex flex-wrap gap-2.5">
          <button
            id="download-json-btn"
            onClick={handleDownloadJson}
            disabled={filteredData.length === 0}
            className="flex items-center gap-2 px-3.5 py-2 text-xs font-semibold bg-blue-600 dark:bg-blue-500 hover:bg-blue-700 dark:hover:bg-blue-600 disabled:bg-slate-200 dark:disabled:bg-slate-700 text-white disabled:text-slate-400 dark:disabled:text-slate-500 rounded-md transition-colors shadow-sm cursor-pointer"
          >
            <Download className="w-4 h-4" />
            {t.btnExportJson}
          </button>
        </div>
      </div>

      {/* Control filters bar */}
      <div className="grid grid-cols-1 md:grid-cols-4 gap-3 bg-[#F8FAFC] dark:bg-slate-900/40 p-3 rounded-md border border-slate-200 dark:border-slate-700">
        {/* Search */}
        <div className="relative md:col-span-2">
          <Search className="w-4 h-4 text-slate-400 dark:text-slate-500 absolute left-3 top-1/2 -translate-y-1/2" />
          <input
            type="text"
            placeholder={t.tableSearchPlaceholder}
            value={searchTerm}
            onChange={(e) => setSearchTerm(e.target.value)}
            className="w-full text-xs pl-9 pr-4 py-2 border border-slate-200 dark:border-slate-700 rounded-md bg-white dark:bg-slate-800 text-slate-700 dark:text-slate-100 focus:outline-none focus:ring-2 focus:ring-blue-500/10 focus:border-blue-500"
          />
        </div>

        {/* Filter Type */}
        <div className="flex items-center gap-2 bg-white dark:bg-slate-800 px-2 py-1 border border-slate-200 dark:border-slate-700 rounded-md">
          <span className="text-[10px] font-bold text-slate-400 dark:text-slate-500 uppercase tracking-wider px-1">{t.filterTypeLabel}</span>
          <select
            value={typeFilter}
            onChange={(e) => setTypeFilter(e.target.value as any)}
            className="text-xs text-slate-700 dark:text-slate-200 font-medium bg-transparent border-none flex-1 focus:outline-none focus:ring-0 cursor-pointer"
          >
            <option value="Todos" className="dark:bg-slate-800">{t.filterAll}</option>
            <option value="Archivo" className="dark:bg-slate-800">{language === "es" ? "Archivos" : "Files"}</option>
            <option value="Carpeta" className="dark:bg-slate-800">{language === "es" ? "Carpetas" : "Folders"}</option>
          </select>
        </div>

        {/* Filter Library */}
        <div className="flex items-center gap-2 bg-white dark:bg-slate-800 px-2 py-1 border border-slate-200 dark:border-slate-700 rounded-md">
          <span className="text-[10px] font-bold text-slate-400 dark:text-slate-500 uppercase tracking-wider px-1">{t.filterLibraryLabel}</span>
          <select
            value={libraryFilter}
            onChange={(e) => setLibraryFilter(e.target.value)}
            className="text-xs text-slate-700 dark:text-slate-200 font-medium bg-transparent border-none flex-1 focus:outline-none focus:ring-0 cursor-pointer truncate"
          >
            {uniqueLibraries.map((lib) => (
              <option key={lib} value={lib} className="dark:bg-slate-800">
                {lib === "Todas" ? t.filterLibraryAll : lib}
              </option>
            ))}
          </select>
        </div>
      </div>

      {/* Grid count and status */}
      <div className="flex items-center justify-between text-xs text-slate-500 dark:text-slate-400 px-1 font-medium">
        <span className="flex items-center gap-1.5">
          {t.filteredStats} <strong className="text-slate-800 dark:text-slate-200">{filteredData.length}</strong> {t.filteredOf} <strong className="text-slate-600 dark:text-slate-400">{data.length}</strong> {t.filteredTotalElements}
        </span>
        <span className="text-slate-400 dark:text-slate-500 text-[11px] flex items-center gap-1">
          <HelpCircle className="w-3.5 h-3.5" />
          {t.calculatedColHeader} <span className="font-mono text-blue-600 dark:text-blue-400 font-semibold bg-blue-50/50 dark:bg-blue-950/40 px-1 py-0.5 rounded">Estructura actual SharePoint</span>
        </span>
      </div>

      {/* Main spreadsheet like table */}
      <div className="border border-slate-200 dark:border-slate-700 rounded-lg overflow-x-auto overflow-y-auto max-h-[550px] shadow-xs">
        <table className="w-full text-left border-collapse min-w-[1150px]">
          <thead>
            <tr className="bg-slate-100/80 dark:bg-slate-900 border-b border-slate-200 dark:border-slate-700 text-xs text-slate-500 dark:text-slate-400 font-bold uppercase tracking-wider select-none sticky top-0 z-10">
              <th className="px-4 py-3 font-semibold bg-slate-100/90 dark:bg-slate-900 w-[12%] whitespace-nowrap">{t.colLibrary}</th>
              <th className="px-4 py-3 font-semibold bg-slate-100/90 dark:bg-slate-900 w-[20%] whitespace-nowrap">{t.colName}</th>
              <th className="px-4 py-3 font-semibold bg-slate-100/90 dark:bg-slate-900 w-[8%] whitespace-nowrap">{t.colType}</th>
              <th className="px-4 py-3 font-semibold bg-slate-100/90 dark:bg-slate-900 w-[32%] whitespace-nowrap min-w-[350px]">{t.colCalculatedPath}</th>
              <th className="px-4 py-3 font-semibold bg-slate-100/90 dark:bg-slate-900 w-[18%] whitespace-nowrap">{t.colMetadata}</th>
              <th className="px-4 py-3 font-semibold bg-slate-100/90 dark:bg-slate-900 w-[6%] text-right whitespace-nowrap">{t.colSize}</th>
              <th className="px-4 py-3 font-semibold bg-slate-100/90 dark:bg-slate-900 w-[6%] text-right whitespace-nowrap">{t.colModified}</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-slate-100 dark:divide-slate-700 text-xs text-slate-600 dark:text-slate-300">
            {filteredData.map((item) => (
              <tr key={item.id} className="hover:bg-slate-50/50 dark:hover:bg-slate-700/20 transition-colors group">
                {/* Library */}
                <td className="px-4 py-3 text-slate-500 dark:text-slate-400 font-medium truncate max-w-[120px]" title={item.library}>
                  {item.library}
                </td>

                {/* Name */}
                <td className="px-4 py-3 font-semibold text-slate-800 dark:text-slate-100">
                  <div className="flex items-center gap-2 max-w-[200px]">
                    {item.type === "Carpeta" ? (
                      <Folder className="w-4 h-4 text-amber-500 shrink-0" />
                    ) : (
                      <File className="w-4 h-4 text-blue-500 shrink-0" />
                    )}
                    <span className="truncate" title={item.name}>{item.name}</span>
                  </div>
                </td>

                {/* Type */}
                <td className="px-4 py-3 shrink-0">
                  <span
                    className={`inline-flex px-1.5 py-0.5 rounded text-[10px] font-bold ${
                      item.type === "Carpeta"
                        ? "bg-amber-50 dark:bg-amber-950/20 text-amber-700 dark:text-amber-400 border border-amber-100 dark:border-amber-900/30"
                        : "bg-blue-50 dark:bg-blue-950/20 text-blue-700 dark:text-blue-400 border border-blue-100 dark:border-blue-900/30"
                    }`}
                  >
                    {item.type === "Carpeta" ? (language === "es" ? "Carpeta" : "Folder") : (language === "es" ? "Archivo" : "File")}
                  </span>
                </td>

                {/* Calculated Column: Estructura actual SharePoint */}
                <td className="px-4 py-3 whitespace-nowrap min-w-[350px]">
                  <div className="flex items-center justify-between gap-3 bg-blue-50/30 dark:bg-blue-950/10 group-hover:bg-blue-50/65 dark:group-hover:bg-blue-950/20 px-2.5 py-1.5 rounded border border-blue-100/50 dark:border-blue-900/40 max-w-[400px] w-full">
                    <div className="flex items-center gap-1 overflow-x-auto whitespace-nowrap scrollbar-none py-0.5 max-w-[330px]">
                      {item["Estructura actual SharePoint"]
                        .split("/")
                        .filter(Boolean)
                        .map((part, idx, arr) => (
                          <React.Fragment key={idx}>
                            {idx > 0 && <span className="text-slate-400 dark:text-slate-500 font-normal text-[10px]">/</span>}
                            <span
                              className={`px-1.5 py-0.5 rounded text-[10px] font-mono font-medium shrink-0 ${
                                idx === arr.length - 1
                                  ? "bg-blue-100 dark:bg-blue-900 text-blue-800 dark:text-blue-200 font-bold"
                                  : "bg-slate-100/80 dark:bg-slate-700 text-slate-600 dark:text-slate-400"
                              }`}
                              title={part}
                            >
                              {part}
                            </span>
                          </React.Fragment>
                        ))}
                    </div>
                    <button
                      onClick={() => handleCopyPath(item["Estructura actual SharePoint"], item.id)}
                      className="text-slate-400 hover:text-blue-600 dark:hover:text-blue-400 transition-colors shrink-0 opacity-0 group-hover:opacity-100 focus:opacity-100 py-1 px-1 bg-white dark:bg-slate-800 border border-slate-200 dark:border-slate-700 rounded cursor-pointer"
                      title={t.copyPathTooltip}
                    >
                      {copiedId === item.id ? (
                        <Check className="w-3.5 h-3.5 text-emerald-500" />
                      ) : (
                        <Copy className="w-3.5 h-3.5" />
                      )}
                    </button>
                  </div>
                </td>

                {/* Custom Metadata */}
                <td className="px-4 py-3">
                  <div className="max-h-[80px] overflow-y-auto pr-1 flex flex-col gap-1 scrollbar-thin max-w-[220px]">
                    {Object.keys(item.customMetadata).length > 0 ? (
                      Object.entries(item.customMetadata).map(([key, val]) => (
                        <div
                          key={key}
                          className="px-2 py-0.5 bg-slate-50 dark:bg-slate-900 border border-slate-150 dark:border-slate-700 text-slate-600 dark:text-slate-400 text-[10px] font-mono rounded flex items-center justify-between gap-1 break-all"
                          title={`${key}: ${val}`}
                        >
                          <span className="font-semibold text-slate-700 dark:text-slate-400 shrink-0">{key}:</span>
                          <span className="text-slate-900 dark:text-slate-200 truncate">{val}</span>
                        </div>
                      ))
                    ) : (
                      <span className="text-slate-400 dark:text-slate-500 text-[10px] italic">{t.noMetadataText}</span>
                    )}
                  </div>
                </td>

                {/* Size */}
                <td className="px-4 py-3 text-right font-mono text-[11px] text-slate-500 dark:text-slate-400">
                  {formatSize(item.size)}
                </td>

                {/* Last Modified */}
                <td className="px-4 py-3 text-right text-slate-400 dark:text-slate-500 whitespace-nowrap">
                  {formatDate(item.lastModified)}
                </td>
              </tr>
            ))}

            {filteredData.length === 0 && (
              <tr>
                <td colSpan={7} className="px-6 py-12 text-center text-slate-400 dark:text-slate-500">
                  <div className="flex flex-col items-center justify-center gap-1.5">
                    <Search className="w-8 h-8 text-slate-300 dark:text-slate-600" />
                    <span className="text-xs font-semibold text-slate-600 dark:text-slate-400">{t.noResultsTitle}</span>
                    <span className="text-xs">{t.noResultsDesc}</span>
                  </div>
                </td>
              </tr>
            )}
          </tbody>
        </table>
      </div>
    </div>
  );
}

