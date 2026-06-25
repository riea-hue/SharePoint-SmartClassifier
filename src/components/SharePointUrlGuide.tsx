import React, { useState } from "react";
import { HelpCircle, ChevronDown, ChevronUp, Link, AlertTriangle, CheckCircle } from "lucide-react";
import { Language, translations } from "../translations";

interface SharePointUrlGuideProps {
  language: Language;
}

export default function SharePointUrlGuide({ language }: SharePointUrlGuideProps) {
  const [isOpen, setIsOpen] = useState(true);
  const t = translations[language];

  return (
    <div id="sharepoint-url-guide-container" className="bg-white dark:bg-slate-800/90 border border-slate-200 dark:border-slate-700 rounded-xl overflow-hidden shadow-sm transition-all duration-300">
      <button
        id="toggle-url-guide-btn"
        onClick={() => setIsOpen(!isOpen)}
        className="w-full flex items-center justify-between p-5 hover:bg-slate-50/80 dark:hover:bg-slate-700/50 transition-colors text-left cursor-pointer"
      >
        <div className="flex items-center gap-3">
          <div className="p-2 bg-emerald-50 dark:bg-emerald-950/40 text-emerald-600 dark:text-emerald-400 rounded-lg">
            <HelpCircle className="w-5 h-5" />
          </div>
          <div>
            <h3 className="font-semibold text-slate-800 dark:text-slate-100 text-base">{t.guideTitle || "Guía de URL de SharePoint Site"}</h3>
            <p className="text-xs text-slate-500 dark:text-slate-400 mt-0.5">{t.guideSubtitle || "Aprende cómo introducir el enlace correcto para escanear y clasificar tus archivos"}</p>
          </div>
        </div>
        {isOpen ? <ChevronUp className="w-5 h-5 text-slate-400 dark:text-slate-500" /> : <ChevronDown className="w-5 h-5 text-slate-400 dark:text-slate-500" />}
      </button>

      {isOpen && (
        <div className="p-6 border-t border-slate-100 dark:border-slate-700 bg-slate-50/30 dark:bg-slate-800/20 space-y-5 text-sm text-slate-600 dark:text-slate-300">
          <div className="grid grid-cols-1 md:grid-cols-2 gap-5">
            {/* Correct URL Guide */}
            <div className="p-4 bg-emerald-50/40 dark:bg-emerald-950/20 border border-emerald-100 dark:border-emerald-900/40 rounded-xl space-y-2">
              <div className="flex items-center gap-2 text-emerald-800 dark:text-emerald-300 font-semibold">
                <CheckCircle className="w-4 h-4 text-emerald-600 dark:text-emerald-400 shrink-0" />
                <span>{t.requiredFormatTitle || "Formato de URL Requerido (Ejemplo Correcto)"}</span>
              </div>
              <p className="text-xs text-slate-600 dark:text-slate-300 leading-relaxed">
                {language === "es" ? (
                  <>Debes ingresar obligatoriamente la <strong>URL raíz del sitio de SharePoint</strong>. Este enlace generalmente apunta al nombre del espacio o departamento de tu organización.</>
                ) : (
                  <>You must enter the <strong>root URL of the SharePoint site</strong>. This link typically points to the space or department name of your organization.</>
                )}
              </p>
              <div className="bg-white dark:bg-slate-900/80 p-2.5 rounded border border-emerald-200/60 dark:border-emerald-800/40 font-mono text-xs text-emerald-700 dark:text-emerald-400 break-all">
                https://tuempresa.sharepoint.com/sites/Finanzas
              </div>
              <p className="text-[11px] text-emerald-600 dark:text-emerald-400 font-medium">
                {t.requiredFormatSuccess || "✓ Permite escanear todas las bibliotecas de documentos y carpetas de ese sitio."}
              </p>
            </div>

            {/* Incorrect URL Guide */}
            <div className="p-4 bg-amber-50/40 dark:bg-amber-950/10 border border-amber-100 dark:border-amber-900/30 rounded-xl space-y-2">
              <div className="flex items-center gap-2 text-amber-800 dark:text-amber-300 font-semibold">
                <AlertTriangle className="w-4 h-4 text-amber-600 dark:text-amber-400 shrink-0" />
                <span>{t.whatNotToEnterTitle || "¿Qué NO debes ingresar?"}</span>
              </div>
              <p className="text-xs text-slate-600 dark:text-slate-300 leading-relaxed">
                {language === "es" ? (
                  <><strong>No introduzcas enlaces directos</strong> a carpetas específicas, subcarpetas, documentos o links compartidos desde el explorador web.</>
                ) : (
                  <><strong>Do not enter direct links</strong> to specific folders, subfolders, documents, or shared links from the web browser.</>
                )}
              </p>
              <div className="bg-white dark:bg-slate-900/80 p-2.5 rounded border border-amber-200/60 dark:border-amber-800/30 font-mono text-xs text-amber-700 dark:text-amber-400 line-through opacity-70 break-all">
                https://tuempresa.sharepoint.com/.../AllItems.aspx?id=%2Fsites%2FFinanzas...
              </div>
              <p className="text-[11px] text-amber-700 dark:text-amber-400 font-medium">
                {t.whatNotToEnterError || "✗ Los enlaces profundos o de archivos individuales causarán fallos de conexión."}
              </p>
            </div>
          </div>

          <div className="flex gap-2.5 items-start p-3.5 bg-slate-100/80 dark:bg-slate-800/60 rounded-lg text-xs text-slate-500 dark:text-slate-400">
            <Link className="w-4 h-4 text-slate-400 dark:text-slate-500 mt-0.5 shrink-0" />
            <p className="leading-relaxed">
              <strong>{t.noteTitle || "Nota didáctica:"}</strong> {t.noteText || "Una vez conectado al sitio raíz, la aplicación detectará y mapeará automáticamente todas las carpetas, archivos y bibliotecas de documentos internas, permitiéndote explorarlas libremente."}
            </p>
          </div>
        </div>
      )}
    </div>
  );
}

