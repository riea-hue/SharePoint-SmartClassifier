import express from "express";
import path from "path";
import dotenv from "dotenv";
import { createServer as createViteServer } from "vite";
import * as msal from "@azure/msal-node";
import { GoogleGenAI, Type } from "@google/genai";

// Load environment variables
dotenv.config();

const app = express();
const PORT = 3000;

app.use(express.json());

// Lazy initialization of Gemini Client
let aiInstance: GoogleGenAI | null = null;
function getGeminiClient(): GoogleGenAI {
  if (!aiInstance) {
    const key = process.env.GEMINI_API_KEY;
    if (!key) {
      throw new Error("GEMINI_API_KEY environment variable is required");
    }
    aiInstance = new GoogleGenAI({
      apiKey: key,
      httpOptions: {
        headers: {
          'User-Agent': 'aistudio-build',
        }
      }
    });
  }
  return aiInstance;
}

// MSAL configuration
const msalConfig: msal.Configuration = {
  auth: {
    clientId: process.env.MICROSOFT_CLIENT_ID || "",
    authority: `https://login.microsoftonline.com/${process.env.MICROSOFT_TENANT_ID || "common"}`,
    clientSecret: process.env.MICROSOFT_CLIENT_SECRET || "",
  }
};

// Lazy initialization of MSAL Confidential Client Application
let msalClient: msal.ConfidentialClientApplication | null = null;

function getMsalClient(): msal.ConfidentialClientApplication {
  if (!msalClient) {
    if (!process.env.MICROSOFT_CLIENT_ID || !process.env.MICROSOFT_CLIENT_SECRET) {
      throw new Error("MICROSOFT_CLIENT_ID and MICROSOFT_CLIENT_SECRET environment variables are required");
    }
    msalClient = new msal.ConfidentialClientApplication(msalConfig);
  }
  return msalClient;
}

// Redirect URI
const getRedirectUri = (): string => {
  const appUrl = process.env.APP_URL || "http://localhost:3000";
  // Trim any trailing slash to ensure clean redirect match
  return `${appUrl.replace(/\/$/, "")}/auth/callback`;
};

// Scopes required for the application
const SCOPES = ["User.Read", "Sites.Read.All", "Files.Read.All"];

// Helper to check if credentials are configured
const isCredentialsConfigured = (): boolean => {
  return !!(process.env.MICROSOFT_CLIENT_ID && process.env.MICROSOFT_CLIENT_SECRET);
};

// Mock Data for Demo Mode
const MOCK_USER = {
  displayName: "Riea Administrador RaaS",
  mail: "riea@riea-raas-ia.com",
  userPrincipalName: "riea@riea-raas-ia.com",
  jobTitle: "Senior Cloud Engineer",
};

const MOCK_SHAREPOINT_SITES = [
  {
    id: "site-root-1234",
    name: "SharePoint Intranet Principal",
    displayName: "Intranet Principal RaaS",
    webUrl: "https://riearaas.sharepoint.com/sites/IntranetPrincipal",
    description: "Sitio principal de la compañía para recursos y documentación general."
  },
  {
    id: "site-finanzas-5678",
    name: "Finanzas y Contabilidad",
    displayName: "Finanzas y Contabilidad",
    webUrl: "https://riearaas.sharepoint.com/sites/Finanzas",
    description: "Espacio de trabajo privado para presupuestos, informes financieros y auditorías."
  }
];

const MOCK_DRIVES = {
  "site-root-1234": [
    { id: "drive-general-1", name: "Documentos Compartidos", description: "Biblioteca de documentos generales" },
    { id: "drive-marketing-2", name: "Recursos de Marca y Marketing", description: "Materiales y activos de marketing" }
  ],
  "site-finanzas-5678": [
    { id: "drive-auditoria-3", name: "Auditorías Anuales", description: "Informes de auditoría interna y externa" },
    { id: "drive-presupuestos-4", name: "Presupuestos", description: "Histórico y planificación de presupuestos" }
  ]
};

const MOCK_DRIVE_ITEMS: Record<string, any[]> = {
  "drive-general-1": [
    { id: "item-g1", name: "Políticas de Empresa", folder: {}, size: 0, webUrl: "https://...", lastModifiedDateTime: "2026-03-12T14:32:00Z" },
    { id: "item-g2", name: "Manual de Bienvenida.pdf", file: {}, size: 2450000, webUrl: "https://...", lastModifiedDateTime: "2026-04-10T09:15:00Z", metadata: { Estado: "Aprobado", Departamento: "RRHH", Version: "1.0" } },
    { id: "item-g3", name: "Plantilla Presentación RaaS.pptx", file: {}, size: 15400000, webUrl: "https://...", lastModifiedDateTime: "2026-05-20T11:45:00Z", metadata: { Estado: "Borrador", Departamento: "Marketing", Version: "0.8" } }
  ],
  "drive-marketing-2": [
    { id: "item-m1", name: "Logos", folder: {}, size: 0, webUrl: "https://...", lastModifiedDateTime: "2026-01-15T16:22:00Z" },
    { id: "item-m2", name: "Estrategia Q3.docx", file: {}, size: 1200000, webUrl: "https://...", lastModifiedDateTime: "2026-06-01T10:00:00Z", metadata: { Estado: "En Revisión", Departamento: "Marketing", Version: "2.3" } }
  ],
  "drive-auditoria-3": [
    { id: "item-a1", name: "Auditoría 2025.xlsx", file: {}, size: 4800000, webUrl: "https://...", lastModifiedDateTime: "2026-02-18T08:30:00Z", metadata: { Estado: "Aprobado", Departamento: "Auditoría Interna", Version: "3.2" } },
    { id: "item-a2", name: "Requerimientos de Entidades Reguladoras.pdf", file: {}, size: 9500000, webUrl: "https://...", lastModifiedDateTime: "2026-05-02T13:12:00Z", metadata: { Estado: "Vigente", Departamento: "Cumplimiento", Version: "1.1" } }
  ],
  "drive-presupuestos-4": [
    { id: "item-p1", name: "Presupuestos 2026", folder: {}, size: 0, webUrl: "https://...", lastModifiedDateTime: "2026-04-05T17:00:00Z" },
    { id: "item-p2", name: "Control_Gastos_Q1.xlsx", file: {}, size: 3400000, webUrl: "https://...", lastModifiedDateTime: "2026-04-10T11:00:00Z", metadata: { Estado: "Aprobado", Departamento: "Contabilidad" } },
    { id: "item-p3", name: "Control_Gastos_Q2.xlsx", file: {}, size: 4100000, webUrl: "https://...", lastModifiedDateTime: "2026-05-12T14:30:00Z", metadata: { Estado: "En Revisión", Departamento: "Contabilidad" } },
    { id: "item-p4", name: "Reporte_Desviaciones_2025.pdf", file: {}, size: 1200000, webUrl: "https://...", lastModifiedDateTime: "2026-01-20T09:15:00Z", metadata: { Estado: "Archivado", Departamento: "Finanzas" } },
    { id: "item-p5", name: "Plan_Reduccion_Costes.docx", file: {}, size: 850000, webUrl: "https://...", lastModifiedDateTime: "2026-06-02T16:00:00Z", metadata: { Estado: "Borrador", Departamento: "Dirección" } },
    { id: "item-p6", name: "Facturas_Pendientes_Junio.xlsx", file: {}, size: 620000, webUrl: "https://...", lastModifiedDateTime: "2026-06-20T10:00:00Z", metadata: { Estado: "Pendiente", Departamento: "Cuentas por Pagar" } },
    { id: "item-p7", name: "Presupuesto_Sistemas_TI.xlsx", file: {}, size: 1800000, webUrl: "https://...", lastModifiedDateTime: "2026-05-15T12:00:00Z", metadata: { Estado: "Aprobado", Departamento: "Sistemas" } },
    { id: "item-p8", name: "Contrato_Licencias_Microsoft.pdf", file: {}, size: 5400000, webUrl: "https://...", lastModifiedDateTime: "2026-05-22T09:30:00Z", metadata: { Estado: "Firmado", Departamento: "Legal" } },
    { id: "item-p9", name: "Fondo_Fijo_Caja_Chica.xlsx", file: {}, size: 220000, webUrl: "https://...", lastModifiedDateTime: "2026-06-18T15:20:00Z", metadata: { Estado: "Aprobado", Departamento: "Tesorería" } },
    { id: "item-p10", name: "Inversiones_Capital_RaaS_2026.pdf", file: {}, size: 7300000, webUrl: "https://...", lastModifiedDateTime: "2026-03-30T11:40:00Z", metadata: { Estado: "Aprobado", Departamento: "Finanzas" } },
    { id: "item-p11", name: "Auditoria_Proveedores_TI.pdf", file: {}, size: 2800000, webUrl: "https://...", lastModifiedDateTime: "2026-02-25T14:10:00Z", metadata: { Estado: "Aprobado", Departamento: "Auditoría Interna" } },
    { id: "item-p12", name: "Presupuesto_RRHH_Contratacion.xlsx", file: {}, size: 1100000, webUrl: "https://...", lastModifiedDateTime: "2026-04-18T10:50:00Z", metadata: { Estado: "En Revisión", Departamento: "RRHH" } },
    { id: "item-p13", name: "Costos_Logistica_Distribucion.xlsx", file: {}, size: 3900000, webUrl: "https://...", lastModifiedDateTime: "2026-05-08T16:15:00Z", metadata: { Estado: "Aprobado", Departamento: "Operaciones" } },
    { id: "item-p14", name: "Plan_Anual_Incentivos.docx", file: {}, size: 950000, webUrl: "https://...", lastModifiedDateTime: "2026-06-10T13:40:00Z", metadata: { Estado: "Borrador", Departamento: "RRHH" } },
    { id: "item-p15", name: "Politica_Reembolso_Gastos.pdf", file: {}, size: 1500000, webUrl: "https://...", lastModifiedDateTime: "2026-01-10T11:00:00Z", metadata: { Estado: "Vigente", Departamento: "Finanzas" } }
  ]
};

// Mock nested items
const MOCK_CHILD_ITEMS: Record<string, any[]> = {
  "item-g1": [
    { id: "item-g1-sub1", name: "Teletrabajo_Politica.pdf", file: {}, size: 1420000, webUrl: "https://...", lastModifiedDateTime: "2026-02-10T11:20:00Z", metadata: { Estado: "Aprobado", Departamento: "RRHH", Version: "2.0" } },
    { id: "item-g1-sub2", name: "Código de Conducta RaaS.pdf", file: {}, size: 3100000, webUrl: "https://...", lastModifiedDateTime: "2026-01-20T10:00:00Z", metadata: { Estado: "Aprobado", Departamento: "Legal", Version: "1.5" } }
  ],
  "item-m1": [
    { id: "item-m1-sub1", name: "Logo_RaaS_Vertical.png", file: {}, size: 850000, webUrl: "https://...", lastModifiedDateTime: "2026-01-15T16:25:00Z", metadata: { Estado: "Aprobado", Departamento: "Diseño", Version: "1.0" } },
    { id: "item-m1-sub2", name: "Logo_RaaS_Horizontal.svg", file: {}, size: 42000, webUrl: "https://...", lastModifiedDateTime: "2026-01-15T16:30:00Z", metadata: { Estado: "Aprobado", Departamento: "Diseño", Version: "1.0" } }
  ],
  "item-p1": [
    { id: "item-p1-sub1", name: "Plan_Anual_Finanzas.xlsx", file: {}, size: 12500000, webUrl: "https://...", lastModifiedDateTime: "2026-05-28T09:00:00Z", metadata: { Estado: "En Revisión", Departamento: "Finanzas", Version: "0.9" } },
    { id: "item-p1-sub2", name: "Proyecciones_Crecimiento_Q3-Q4.xlsx", file: {}, size: 8200000, webUrl: "https://...", lastModifiedDateTime: "2026-06-15T15:40:00Z", metadata: { Estado: "Borrador", Departamento: "Planificación", Version: "0.4" } }
  ]
};

// API ROUTES

// 1. Get Environment Configuration Status (to show warnings if Azure AD is not configured)
app.get("/api/config", (req, res) => {
  res.json({
    isConfigured: isCredentialsConfigured(),
    redirectUri: getRedirectUri(),
    clientId: process.env.MICROSOFT_CLIENT_ID || "",
    tenantId: process.env.MICROSOFT_TENANT_ID || "common",
    appName: "Microsoft Graph & SharePoint Mapper"
  });
});

// 2. Generate MSAL Auth URL
app.get("/api/auth/url", async (req, res) => {
  try {
    const redirectUri = getRedirectUri();
    
    // In case credentials are not configured, send a state indicating demo mode
    if (!isCredentialsConfigured()) {
      return res.status(400).json({ 
        error: "NotConfigured", 
        message: "Las credenciales de Microsoft Azure AD no están configuradas en .env.example. Por favor configúrelas o active el Modo Demo." 
      });
    }

    const cca = getMsalClient();
    const authCodeUrlParameters = {
      scopes: SCOPES,
      redirectUri: redirectUri,
    };

    const authUrl = await cca.getAuthCodeUrl(authCodeUrlParameters);
    res.json({ url: authUrl });
  } catch (error: any) {
    console.error("Error generating auth URL:", error);
    res.status(500).json({ error: "AuthUrlError", message: error.message });
  }
});

// 3. OAuth Callback handler
app.get(["/auth/callback", "/auth/callback/"], async (req, res) => {
  const { code, error, error_description } = req.query;

  if (error) {
    console.error("Auth Callback error:", error, error_description);
    return res.send(`
      <html>
        <body style="font-family: sans-serif; display: flex; flex-direction: column; align-items: center; justify-content: center; height: 100vh; background-color: #fef2f2; color: #991b1b; padding: 24px;">
          <div style="max-width: 500px; text-align: center; border: 1px solid #fca5a5; padding: 32px; border-radius: 8px; background: white; box-shadow: 0 4px 6px -1px rgb(0 0 0 / 0.1);">
            <h2 style="margin-top: 0;">Error de Autenticación</h2>
            <p><strong>Código de error:</strong> ${error}</p>
            <p>${error_description || "Ocurrió un error al iniciar sesión con Microsoft."}</p>
            <button onclick="window.close()" style="background-color: #dc2626; color: white; border: none; padding: 10px 20px; border-radius: 4px; cursor: pointer; font-weight: bold; margin-top: 16px;">Cerrar Ventana</button>
          </div>
        </body>
      </html>
    `);
  }

  if (!code) {
    return res.status(400).send("No authorization code provided");
  }

  try {
    const cca = getMsalClient();
    const redirectUri = getRedirectUri();
    
    const tokenRequest = {
      code: code as string,
      scopes: SCOPES,
      redirectUri: redirectUri,
    };

    const response = await cca.acquireTokenByCode(tokenRequest);
    const accessToken = response.accessToken;
    
    // Call Graph /me to get official profile details
    let userProfile = {
      displayName: response.account?.name || "Usuario de SharePoint",
      mail: response.account?.username || "usuario@sharepoint.com",
      userPrincipalName: response.account?.username || "usuario@sharepoint.com",
      jobTitle: "Miembro del Equipo"
    };

    try {
      const meResponse = await fetch("https://graph.microsoft.com/v1.0/me", {
        headers: {
          Authorization: `Bearer ${accessToken}`
        }
      });
      if (meResponse.ok) {
        const meData = await meResponse.json();
        userProfile = {
          displayName: meData.displayName || userProfile.displayName,
          mail: meData.mail || meData.userPrincipalName || userProfile.mail,
          userPrincipalName: meData.userPrincipalName || userProfile.userPrincipalName,
          jobTitle: meData.jobTitle || userProfile.jobTitle
        };
      }
    } catch (e) {
      console.warn("Could not fetch extended user profile:", e);
    }

    // PREPARACIÓN DE REGISTRO INVISIBLE Y CONTROL FREEMIUM (Paso 4):
    // Captura del correo electrónico del usuario desde el perfil de Microsoft (ya almacenado en userProfile.mail).
    // Estructura lista para conectar una base de datos (por ejemplo, PostgreSQL o Firestore) 
    // que asocie este correo a un límite de 1 uso gratuito de escaneo, bloqueando futuros escaneos si el correo ya consumió su prueba.
    /*
    try {
      const userEmail = userProfile.mail;
      const db = await getDatabaseConnection(); // Conexión a la base de datos relacional / NoSQL
      
      // Verificar si el usuario ya existe en la base de datos de control freemium
      const usageRecord = await db.query("SELECT scan_count, trial_consumed FROM user_limits WHERE email = $1", [userEmail]);
      
      if (usageRecord.rows.length > 0) {
        const { scan_count, trial_consumed } = usageRecord.rows[0];
        if (trial_consumed || scan_count >= 1) {
          console.log(`Límite freemium alcanzado para: ${userEmail}. Escaneos previos: ${scan_count}`);
          // Bloquear escaneo y solicitar actualización a Pro:
          // throw new Error("Ya has consumido tu escaneo de prueba gratuito. Adquiere la versión Pro.");
        }
      } else {
        // Registrar al usuario de forma invisible al iniciar sesión por primera vez con 0 escaneos consumidos
        await db.query("INSERT INTO user_limits (email, scan_count, trial_consumed) VALUES ($1, 0, false)", [userEmail]);
        console.log(`Usuario nuevo registrado de forma invisible: ${userEmail}`);
      }
    } catch (dbError) {
      console.error("Error en el control invisible de base de datos freemium:", dbError);
    }
    */

    // Send success message and data back to parent window
    res.send(`
      <html>
        <body style="font-family: sans-serif; display: flex; flex-direction: column; align-items: center; justify-content: center; height: 100vh; background-color: #f0fdf4; color: #166534; padding: 24px;">
          <div style="max-width: 500px; text-align: center; border: 1px solid #bbf7d0; padding: 32px; border-radius: 8px; background: white; box-shadow: 0 4px 6px -1px rgb(0 0 0 / 0.1);">
            <svg style="width: 64px; height: 64px; color: #22c55e; margin: 0 auto 16px;" fill="none" stroke="currentColor" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg">
              <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z"></path>
            </svg>
            <h2 style="margin-top: 0; color: #15803d;">¡Autenticación Exitosa!</h2>
            <p>Se ha conectado con éxito como <strong>${userProfile.mail}</strong>.</p>
            <p style="color: #666; font-size: 14px;">Iniciando sesión... Puedes cerrar esta pestaña y volver a la aplicación si no se cierra sola.</p>
            <button onclick="window.close()" style="margin-top: 16px; padding: 10px 20px; background-color: #16a34a; color: white; border: none; border-radius: 6px; cursor: pointer; font-weight: 600; font-size: 13px;">Cerrar esta pestaña y volver</button>
            <script>
              try {
                localStorage.setItem('microsoft_access_token', '${accessToken}');
                localStorage.setItem('microsoft_user_profile', JSON.stringify(${JSON.stringify(userProfile)}));
              } catch (e) {
                console.error("No se pudo guardar la sesión en localStorage:", e);
              }

              // Notify the parent frame if opener exists
              if (window.opener) {
                try {
                  window.opener.postMessage({ 
                    type: 'OAUTH_AUTH_SUCCESS', 
                    token: '${accessToken}',
                    user: ${JSON.stringify(userProfile)}
                  }, '*');
                } catch (e) {
                  console.error("Error sending postMessage:", e);
                }
              }

              // Try to close window automatically
              setTimeout(() => {
                try {
                  window.close();
                } catch (e) {
                  console.warn("Could not auto-close window:", e);
                }
              }, 1500);
            </script>
          </div>
        </body>
      </html>
    `);
  } catch (error: any) {
    console.error("Error exchanging code for token:", error);
    res.status(500).send(`
      <html>
        <body style="font-family: sans-serif; display: flex; flex-direction: column; align-items: center; justify-content: center; height: 100vh; background-color: #fef2f2; color: #991b1b; padding: 24px;">
          <div style="max-width: 500px; text-align: center; border: 1px solid #fca5a5; padding: 32px; border-radius: 8px; background: white; box-shadow: 0 4px 6px -1px rgb(0 0 0 / 0.1);">
            <h2 style="margin-top: 0;">Falla de Intercambio de Tokens</h2>
            <p>${error.message || "No se pudo canjear el código por un token de acceso de Microsoft."}</p>
            <button onclick="window.close()" style="background-color: #dc2626; color: white; border: none; padding: 10px 20px; border-radius: 4px; cursor: pointer; font-weight: bold; margin-top: 16px;">Cerrar Ventana</button>
          </div>
        </body>
      </html>
    `);
  }
});

// 4. Resolve SharePoint Site URL to Site ID & Check Permissions (Access Validation)
app.post("/api/sharepoint/resolve-site", async (req, res) => {
  const { siteUrl, isDemo } = req.body;
  const authHeader = req.headers.authorization;

  if (isDemo || !authHeader) {
    // Demo Mode fallback
    if (!siteUrl) {
      return res.json({ site: MOCK_SHAREPOINT_SITES[0] });
    }
    // Simple site selection mock
    const foundSite = MOCK_SHAREPOINT_SITES.find(
      s => s.webUrl.toLowerCase().includes(siteUrl.toLowerCase()) || s.name.toLowerCase().includes(siteUrl.toLowerCase())
    );

    if (!foundSite) {
      return res.status(403).json({
        error: "AccessDenied",
        message: "Acceso denegado: No tienes permisos para acceder a este sitio de SharePoint o el sitio no existe en el catálogo."
      });
    }

    return res.json({ site: foundSite });
  }

  // Real Mode Graph API call
  try {
    const token = authHeader.split(" ")[1];
    
    // Parse the site url to fetch from graph.
    // E.g., URL format: https://yourtenant.sharepoint.com/sites/yoursitename
    let graphUrl = "https://graph.microsoft.com/v1.0/sites/root"; // default to root site

    if (siteUrl && siteUrl.trim() !== "") {
      try {
        const parsedUrl = new URL(siteUrl);
        const hostname = parsedUrl.hostname;
        const sitePath = parsedUrl.pathname.replace(/^\/$/, ""); // clear slash
        
        if (sitePath && sitePath !== "/") {
          graphUrl = `https://graph.microsoft.com/v1.0/sites/${hostname}:${sitePath}`;
        } else {
          graphUrl = `https://graph.microsoft.com/v1.0/sites/${hostname}`;
        }
      } catch (err) {
        // If it's not a URL, check if it's an absolute site ID, otherwise treat as relative path on root domain
        if (siteUrl.includes(",") || siteUrl.length > 20) {
          graphUrl = `https://graph.microsoft.com/v1.0/sites/${siteUrl}`;
        } else {
          return res.status(400).json({
            error: "InvalidUrl",
            message: "Por favor proporciona un URL de sitio válido (ej: https://tenant.sharepoint.com/sites/nombre_sitio)."
          });
        }
      }
    }

    const response = await fetch(graphUrl, {
      headers: {
        Authorization: `Bearer ${token}`
      }
    });

    if (!response.ok) {
      if (response.status === 403) {
        return res.status(403).json({
          error: "AccessDenied",
          message: "Acceso denegado: Su cuenta no cuenta con permisos de lectura (Sites.Read.All) para este sitio de SharePoint."
        });
      } else if (response.status === 401) {
        return res.status(401).json({
          error: "TokenExpired",
          message: "Su sesión de Microsoft Graph ha expirado. Por favor inicie sesión nuevamente."
        });
      } else {
        const errText = await response.text();
        return res.status(response.status).json({
          error: "GraphError",
          message: `Microsoft Graph API retornó error ${response.status}: ${errText}`
        });
      }
    }

    const siteData = await response.json();
    res.json({
      site: {
        id: siteData.id,
        name: siteData.name,
        displayName: siteData.displayName || siteData.name,
        webUrl: siteData.webUrl,
        description: siteData.description || ""
      }
    });
  } catch (error: any) {
    console.error("Error resolving SharePoint site:", error);
    res.status(500).json({ error: "ServerError", message: error.message });
  }
});

// 5. Fetch Document Libraries (Drives) for a resolved SharePoint Site
app.get("/api/sharepoint/sites/:siteId/drives", async (req, res) => {
  const { siteId } = req.params;
  const { isDemo } = req.query;
  const authHeader = req.headers.authorization;

  if (isDemo === "true" || !authHeader) {
    const drives = MOCK_DRIVES[siteId as keyof typeof MOCK_DRIVES] || MOCK_DRIVES["site-root-1234"];
    return res.json({ drives });
  }

  try {
    const token = authHeader.split(" ")[1];
    const graphUrl = `https://graph.microsoft.com/v1.0/sites/${siteId}/drives`;

    const response = await fetch(graphUrl, {
      headers: {
        Authorization: `Bearer ${token}`
      }
    });

    if (!response.ok) {
      return res.status(response.status).json({
        error: "GraphError",
        message: `No se pudieron cargar las bibliotecas de documentos: ${response.statusText}`
      });
    }

    const data = await response.json();
    res.json({
      drives: data.value.map((d: any) => ({
        id: d.id,
        name: d.name,
        description: d.description || ""
      }))
    });
  } catch (error: any) {
    console.error("Error fetching drives:", error);
    res.status(500).json({ error: "ServerError", message: error.message });
  }
});

// 6. Fetch Items (Files and Folders) in a specific directory/folder of a Drive
app.get("/api/sharepoint/drives/:driveId/items", async (req, res) => {
  const { driveId } = req.params;
  const { folderId, isDemo } = req.query;
  const authHeader = req.headers.authorization;

  if (isDemo === "true" || !authHeader) {
    // If folderId is provided, look in nested mock child items
    if (folderId) {
      const items = MOCK_CHILD_ITEMS[folderId as string] || [];
      return res.json({ items });
    }
    const items = MOCK_DRIVE_ITEMS[driveId as string] || [];
    return res.json({ items });
  }

  try {
    const token = authHeader.split(" ")[1];
    
    // If folderId is provided, get items of that folder. Otherwise get root.
    const graphUrl = folderId 
      ? `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${folderId}/children?$expand=listItem`
      : `https://graph.microsoft.com/v1.0/drives/${driveId}/root/children?$expand=listItem`;

    const response = await fetch(graphUrl, {
      headers: {
        Authorization: `Bearer ${token}`
      }
    });

    if (!response.ok) {
      return res.status(response.status).json({
        error: "GraphError",
        message: `No se pudieron cargar los elementos: ${response.statusText}`
      });
    }

    const data = await response.json();
    
    // Map Microsoft Graph driveItems into a unified item format
    // Extract metadata from listItem/fields if available
    const items = data.value.map((item: any) => {
      const customMetadata: Record<string, string> = {};
      
      // Look for custom columns in the listItem expansion
      if (item.listItem && item.listItem.fields) {
        const fields = item.listItem.fields;
        // Grab fields that are custom (avoiding standard system fields)
        Object.keys(fields).forEach(key => {
          if (!["id", "ContentType", "Created", "Modified", "AuthorLookupId", "EditorLookupId", "_UId", "Edit", "LinkFilenameNoMenu", "LinkFilename", "DocIcon", "FileSizeDisplay", "ItemChildCount", "FolderChildCount"].includes(key)) {
            if (key === "AITags") {
              customMetadata["AI Tags"] = String(fields[key]);
            } else {
              customMetadata[key] = String(fields[key]);
            }
          }
        });
      }

      return {
        id: item.id,
        name: item.name,
        folder: item.folder ? {} : undefined,
        file: item.file ? {} : undefined,
        size: item.size || 0,
        webUrl: item.webUrl,
        lastModifiedDateTime: item.lastModifiedDateTime,
        metadata: Object.keys(customMetadata).length > 0 ? customMetadata : undefined
      };
    });

    res.json({ items });
  } catch (error: any) {
    console.error("Error fetching items:", error);
    res.status(500).json({ error: "ServerError", message: error.message });
  }
});

// 7. Recursive Mapper - Unifies SharePoint structure and returns the full hierarchy in a flat array
// with the calculated column "Estructura actual SharePoint"
app.post("/api/sharepoint/map-structure", async (req, res) => {
  const { siteId, isDemo } = req.body;
  const authHeader = req.headers.authorization;

  if (isDemo || !authHeader) {
    // Construct mapped structure for Mock data
    const mappedItems: any[] = [];
    const drives = MOCK_DRIVES[siteId as keyof typeof MOCK_DRIVES] || MOCK_DRIVES["site-root-1234"];

    for (const d of drives) {
      const rootItems = MOCK_DRIVE_ITEMS[d.id] || [];
      
      for (const item of rootItems) {
        const rootPath = `${d.name}/${item.name}`;
        
        mappedItems.push({
          id: item.id,
          driveId: d.id,
          name: item.name,
          library: d.name,
          type: item.folder ? "Carpeta" : "Archivo",
          size: item.size,
          lastModified: item.lastModifiedDateTime,
          customMetadata: item.metadata || {},
          "Estructura actual SharePoint": rootPath
        });

        // If it's a folder, search inside mock nested items
        if (item.folder) {
          const subItems = MOCK_CHILD_ITEMS[item.id] || [];
          for (const subItem of subItems) {
            const subPath = `${rootPath}/${subItem.name}`;
            mappedItems.push({
              id: subItem.id,
              driveId: d.id,
              name: subItem.name,
              library: d.name,
              type: subItem.folder ? "Carpeta" : "Archivo",
              size: subItem.size,
              lastModified: subItem.lastModifiedDateTime,
              customMetadata: subItem.metadata || {},
              "Estructura actual SharePoint": subPath
            });
          }
        }
      }
    }

    return res.json({ mappedItems });
  }

  // Real Mode Mapping using Graph API
  try {
    const token = authHeader.split(" ")[1];
    
    // First, list all drives
    const drivesRes = await fetch(`https://graph.microsoft.com/v1.0/sites/${siteId}/drives`, {
      headers: { Authorization: `Bearer ${token}` }
    });

    if (!drivesRes.ok) {
      return res.status(drivesRes.status).json({
        error: "GraphError",
        message: `Error al listar bibliotecas de documentos para el mapa: ${drivesRes.statusText}`
      });
    }

    const drivesData = await drivesRes.json();
    const drives = drivesData.value;
    const mappedItems: any[] = [];

    // Helper to recursively fetch child items of a folder
    const traverseFolder = async (driveId: string, driveName: string, folderId: string | null, parentPath: string, depth: number) => {
      // Limit recursion depth to 4 levels to prevent rate limits or long response times
      if (depth > 4) return;

      const graphUrl = folderId
        ? `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${folderId}/children?$expand=listItem`
        : `https://graph.microsoft.com/v1.0/drives/${driveId}/root/children?$expand=listItem`;

      try {
        const response = await fetch(graphUrl, {
          headers: { Authorization: `Bearer ${token}` }
        });

        if (!response.ok) return;

        const data = await response.json();
        for (const item of data.value) {
          const currentPath = parentPath ? `${parentPath}/${item.name}` : `${driveName}/${item.name}`;
          
          // Custom columns
          const customMetadata: Record<string, string> = {};
          if (item.listItem && item.listItem.fields) {
            const fields = item.listItem.fields;
            Object.keys(fields).forEach(key => {
              if (!["id", "ContentType", "Created", "Modified", "AuthorLookupId", "EditorLookupId", "_UId", "Edit", "LinkFilenameNoMenu", "LinkFilename", "DocIcon", "FileSizeDisplay", "ItemChildCount", "FolderChildCount"].includes(key)) {
                if (key === "AITags") {
                  customMetadata["AI Tags"] = String(fields[key]);
                } else {
                  customMetadata[key] = String(fields[key]);
                }
              }
            });
          }

          mappedItems.push({
            id: item.id,
            driveId: driveId,
            name: item.name,
            library: driveName,
            type: item.folder ? "Carpeta" : "Archivo",
            size: item.size || 0,
            lastModified: item.lastModifiedDateTime,
            customMetadata,
            "Estructura actual SharePoint": currentPath
          });

          if (item.folder) {
            // Traversal of subdirectory
            await traverseFolder(driveId, driveName, item.id, currentPath, depth + 1);
          }
        }
      } catch (e) {
        console.error(`Error traversing folder ${folderId} in drive ${driveId}:`, e);
      }
    };

    // Traverse all drives
    for (const drive of drives) {
      await traverseFolder(drive.id, drive.name, null, "", 1);
    }

    res.json({ mappedItems });
  } catch (error: any) {
    console.error("Error mapping structure:", error);
    res.status(500).json({ error: "ServerError", message: error.message });
  }
});

// Session tracker for AI requests to prevent abusive costs (Max 100 per session)
const classificationCounts = new Map<string, number>();

// 8. Sugerir Categorías con Gemini
app.post("/api/gemini/suggest-categories", async (req, res) => {
  const { fileName, library, type, customMetadata, path: itemPath } = req.body;
  const sessionId = req.headers["x-session-id"] || "default-session";
  
  if (!fileName) {
    return res.status(400).json({ error: "MissingFileName", message: "Falta el nombre del archivo para generar categorías." });
  }

  const currentCount = classificationCounts.get(sessionId as string) || 0;
  if (currentCount >= 100) {
    return res.status(429).json({
      error: "QuotaExceeded",
      message: "Límite de procesamiento de IA alcanzado por seguridad del servidor. Si necesitas procesar volúmenes masivos, contacta a RIEA para una solución dedicada corporativa."
    });
  }

  // Increment session counter
  const newCount = currentCount + 1;
  classificationCounts.set(sessionId as string, newCount);

  try {
    const ai = getGeminiClient();
    const prompt = `Analiza el siguiente elemento (archivo o carpeta) del sitio de SharePoint y sugiere exactamente 5 hashtags o etiquetas que representen mejor su contenido, departamento, área, propósito o tipo de documento.
    
Nombre del elemento: "${fileName}"
Ruta en SharePoint: "${itemPath || ''}"
Biblioteca de origen: "${library || 'Documentos'}"
Tipo: "${type || 'Archivo'}"
Metadatos adicionales: ${customMetadata ? JSON.stringify(customMetadata) : 'Ninguno'}

Reglas:
1. Genera exactamente 5 hashtags relevantes en español.
2. Cada hashtag DEBE comenzar con el símbolo '#' (ejemplo: #Contrato, #Finanzas, #Legal, #Auditoria2026, #Proveedor).
3. Devuelve los hashtags como un arreglo JSON de cadenas de texto. No uses formato markdown alrededor, solo el JSON que coincida con el esquema solicitado.`;

    const response = await ai.models.generateContent({
      model: "gemini-3.5-flash",
      contents: prompt,
      config: {
        responseMimeType: "application/json",
        responseSchema: {
          type: Type.ARRAY,
          items: {
            type: Type.STRING
          },
          description: "Lista de exactamente 5 hashtags descriptivos que comiencen con #"
        }
      }
    });

    const text = response.text || "[]";
    const hashtags = JSON.parse(text);
    
    // Ensure all start with '#'
    const formattedHashtags = hashtags.map((tag: string) => {
      let trimmed = tag.trim();
      if (!trimmed.startsWith("#")) {
        trimmed = "#" + trimmed;
      }
      // Replace spaces in hashtags if any
      return trimmed.replace(/\s+/g, "");
    }).slice(0, 5);

    // If less than 5 generated, fill up with some fallbacks based on file extension/library
    while (formattedHashtags.length < 5) {
      const fallbackTags = ["#SharePoint", "#Documento", "#General", "#RaaS", "#Clasificado"];
      const nextTag = fallbackTags[formattedHashtags.length] || `#Etiqueta${formattedHashtags.length + 1}`;
      formattedHashtags.push(nextTag);
    }

    res.json({ hashtags: formattedHashtags, sessionCount: newCount });
  } catch (error: any) {
    console.error("Error generating categories with Gemini:", error);
    // Return friendly fallback tags if GEMINI_API_KEY is not set or if there's any API error
    let fallbackTags = ["#SharePoint", "#Documento", "#General", "#Empresa", "#Archivado"];
    // Try to guess a tag based on library or filename
    if (library) {
      fallbackTags[1] = `#${library.replace(/\s+/g, "")}`;
    }
    const ext = fileName.split('.').pop() || '';
    if (ext && ext.length > 1 && ext.length < 5) {
      fallbackTags[2] = `#${ext.toUpperCase()}`;
    }
    res.json({ 
      hashtags: fallbackTags, 
      sessionCount: newCount,
      warning: "Se usaron etiquetas de respaldo debido a un error o falta de configuración de la API de Gemini: " + error.message 
    });
  }
});

// 9. Actualizar columna 'AI Tags' en SharePoint
app.post("/api/sharepoint/update-metadata", async (req, res) => {
  const { driveId, itemId, metadataValue, isDemo } = req.body;
  const authHeader = req.headers.authorization;

  if (!driveId || !itemId) {
    return res.status(400).json({ error: "MissingParameters", message: "Faltan parámetros requeridos (driveId o itemId)." });
  }

  const tagString = String(metadataValue || "").trim();

  if (isDemo || !authHeader) {
    // Simulated updates to MOCK collections
    let updated = false;

    // Check MOCK_DRIVE_ITEMS
    if (MOCK_DRIVE_ITEMS[driveId]) {
      const item = MOCK_DRIVE_ITEMS[driveId].find(i => i.id === itemId);
      if (item) {
        if (!item.metadata) item.metadata = {};
        item.metadata["AI Tags"] = tagString;
        updated = true;
      }
    }

    // Check MOCK_CHILD_ITEMS
    if (!updated) {
      for (const parentId in MOCK_CHILD_ITEMS) {
        const item = MOCK_CHILD_ITEMS[parentId]?.find(i => i.id === itemId);
        if (item) {
          if (!item.metadata) item.metadata = {};
          item.metadata["AI Tags"] = tagString;
          updated = true;
          break;
        }
      }
    }

    return res.json({ 
      success: true, 
      message: "Metadatos actualizados con éxito en Modo Demo.", 
      updatedFields: { "AI Tags": tagString } 
    });
  }

  // Real Mode Microsoft Graph API call
  try {
    const token = authHeader.split(" ")[1];
    
    // In Microsoft Graph, we update custom columns using the listItem/fields endpoint.
    const graphUrl = `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}/listItem/fields`;
    
    const patchBody = {
      "AITags": tagString
    };

    const response = await fetch(graphUrl, {
      method: "PATCH",
      headers: {
        "Authorization": `Bearer ${token}`,
        "Content-Type": "application/json"
      },
      body: JSON.stringify(patchBody)
    });

    if (!response.ok) {
      const errText = await response.text();
      console.error(`Graph metadata update returned error ${response.status}:`, errText);
      
      // If the field "AI Tags" is not configured yet on SharePoint, it could return 400.
      return res.status(response.status).json({
        error: "GraphUpdateError",
        message: `No se pudo actualizar la columna 'AI Tags' en SharePoint (${response.status}). Asegúrate de que la columna exista con ese nombre exacto. Detalle: ${errText}`
      });
    }

    const resData = await response.json();
    res.json({
      success: true,
      message: "¡La columna 'AI Tags' ha sido actualizada en SharePoint exitosamente!",
      updatedFields: resData
    });
  } catch (error: any) {
    console.error("Error updating SharePoint metadata:", error);
    res.status(500).json({ error: "ServerError", message: error.message });
  }
});


// Serve static frontend files in production, use Vite in development
async function startServer() {
  if (process.env.NODE_ENV !== "production") {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: "spa",
    });
    app.use(vite.middlewares);
  } else {
    const distPath = path.join(process.cwd(), "dist");
    app.use(express.static(distPath));
    app.get("*", (req, res) => {
      res.sendFile(path.join(distPath, "index.html"));
    });
  }

  app.listen(PORT, "0.0.0.0", () => {
    console.log(`Server running on http://localhost:${PORT}`);
  });
}

startServer();
