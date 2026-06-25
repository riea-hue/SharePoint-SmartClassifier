# RIEA SharePoint SmartClassifier v1.1

![React](https://img.shields.io/badge/React-20232A?style=for-the-badge&logo=react&logoColor=61DAFB)
![Node.js](https://img.shields.io/badge/Node.js-339933?style=for-the-badge&logo=nodedotjs&logoColor=white)
![Google Gemini](https://img.shields.io/badge/Google%20Gemini-8E75C2?style=for-the-badge&logo=googlegemini&logoColor=white)
![Microsoft Graph](https://img.shields.io/badge/Microsoft%20Graph-0078D4?style=for-the-badge&logo=microsoft&logoColor=white)
![Tailwind CSS](https://img.shields.io/badge/Tailwind_CSS-38B2AC?style=for-the-badge&logo=tailwind-css&logoColor=white)

> **Clasificación inteligente de documentos y carpetas con Inteligencia Artificial mediante Microsoft Graph API.** Una solución corporativa integral diseñada bajo la marca personal **RIEA** para optimizar la gobernanza de datos y la automatización de procesos en entornos Microsoft 365.

---

## 🚀 Enlace del Proyecto
Probad la aplicación en vivo utilizando el **Modo Demo Interactivo** (sin necesidad de configuraciones ni credenciales de Azure) aquí:
🔗 **[Enlace de tu App en Google Cloud Run]**

---

## 📋 Descripción del Proyecto

En el entorno corporativo actual, las empresas pierden cientos de horas-hombre organizando, etiquetando y buscando archivos dentro de sus repositorios digitales. **SharePoint SmartClassifier** nace para solucionar el caos documental de raíz mediante la intersección de **Procesos, Gobierno de Datos e Inteligencia Artificial**.

La herramienta se conecta de manera segura a un sitio raíz de SharePoint, mapea su estructura jerárquica y utiliza el modelo de lenguaje de última generación **Gemini 1.5 Flash** para analizar el contexto de archivos individuales o carpetas completas. Como resultado, genera y guarda automáticamente **5 hashtags enriquecidos** directamente dentro de una columna de texto múltiple personalizada en SharePoint (`AI Tags`), optimizando de manera drástica el motor de búsqueda interno de la organización.

---

## ✨ Características Principales

### 1. Mapeador de Estructura (Modelo Freemium)
* **Visión Jerárquica:** Mapeo y exploración en tiempo real de las bibliotecas de documentos, carpetas y archivos mediante peticiones eficientes utilizando Microsoft Graph API.
* **Límite de Prueba Estricto:** Restricción automatizada en el backend a los primeros **15 elementos** para cuentas en versión gratuita, acompañado de un banner premium deshabilitado con la leyenda `Coming Soon` para la conversión a la versión Pro.

### 2. Clasificador AI Dinámico (Soporte para Archivos y Carpetas)
* **Análisis Contextual:** Gemini procesa nombres, rutas y extensiones de archivos para deducir de forma precisa la taxonomía ideal. El sistema soporta la clasificación tanto de archivos individuales como de carpetas macro corporativas.
* **Hashtags Editables:** La interfaz despliega los 5 hashtags sugeridos en campos de texto completamente interactivos para que el usuario pueda auditarlos, borrarlos o agregar etiquetas personalizadas antes de consolidarlos.
* **Inyección a SharePoint:** Al confirmar, el sistema concatena los tags separados por espacios y realiza una solicitud `PATCH` a Microsoft Graph actualizando la columna interna `AITags` (utilizando la notación *camelCase* requerida por Microsoft).

### 3. Modo Demo Inteligente (Ideal para Reclutadores / Portfolio)
* **Cero Fricción:** Diseñado específicamente para que los líderes técnicos y *Hiring Managers* prueben la aplicación en 5 segundos sin comprometer la seguridad ni las claves de su organización.
* **Mock AI Context-Aware:** Al activar el interruptor de Modo Demo, el sistema simula un tiempo de carga asíncrono de 1.5s (con un spinner visual) y genera localmente 5 hashtags sumamente coherentes basados en palabras clave del archivo de simulación (ej. archivos con "Factura" disparan tags contables; carpetas disparan tags estructurales).

### 4. Características Premium y Accesibilidad Global
* **Internacionalización (i18n):** Soporte nativo para alternar toda la interfaz de la aplicación de manera instantánea entre **Español e Inglés** mediante un selector en el encabezado.
* **Modo Oscuro (Dark Mode):** Interruptor Sol/Luna con una paleta de colores de alto contraste visual (#0f172a / Slate / Violeta RIEA) que reduce la fatiga visual de los gestores de contenido.
* **Full Viewport Layout:** Renderizado fluido adaptado al 100% de la pantalla (`w-screen h-screen`). Contenedores responsivos controlados con barras de desplazamiento internas para soportar grandes volúmenes de datos sin romper la maquetación.

### 5. FinOps y Portabilidad de Datos (Gobernanza de TI)
* **Control de Costos de API:** Limitación defensiva en el backend de un máximo de **100 llamadas de clasificación por sesión** de usuario para mitigar consumos accidentales de tokens y prevenir ataques de costos. Muestra advertencias de cuota al alcanzar el 90%.
* **Exportador de Reportes:** Botón dedicado para descargar la taxonomía enriquecida por la IA en formatos **CSV** (vista plana para análisis) y **JSON Unificado** (con el árbol completo y todas las columnas extendidas mediante `$expand=fields` de Microsoft Graph).

---

## 🛠️ Arquitectura y Flujo Técnico

La aplicación sigue un enfoque de arquitectura limpia desacoplada:

1. **Frontend (React & Tailwind CSS):** Gestiona de forma interactiva el ciclo de vida de los datos, la autenticación basada en flujos de redirección y la renderización en memoria del árbol documental.
2. **Autenticación (MSAL / OAuth2):** Implementación segura mediante flujos delegados de Microsoft Entra ID utilizando los alcances mínimos de lectura y escritura: `User.Read`, `Sites.ReadWrite.All` y `Files.ReadWrite.All`.
3. **Backend (Node.js/Express):** Actúa como un proxy seguro e intermedio que resguarda la API Key de Google Gemini y gestiona el límite de créditos invisibles por correo electrónico capturado en el perfil de Microsoft del usuario.
4. **Capa IA (Google Gemini 1.5 Flash SDK):** Utiliza esquemas de respuesta JSON estructurados para asegurar que el modelo siempre devuelva un formato de array de strings procesable por el software.

---

## ⚙️ Configuración e Instalación Local

Si deseas clonar este repositorio y ejecutarlo en tu propio entorno de desarrollo, sigue estos pasos:

### 1. Clonar el repositorio
bash
git clone [https://github.com/tu-usuario/RIEA-SharePoint-SmartClassifier.git](https://github.com/tu-usuario/RIEA-SharePoint-SmartClassifier.git)

2. Configurar variables de entorno (.env)
Crea un archivo .env en la raíz del proyecto con los siguientes parámetros:

Fragmento de código
PORT=5000
GEMINI_API_KEY=tu_gemini_api_key_aqui
MICROSOFT_CLIENT_ID=tu_client_id_de_azure_ad
MICROSOFT_TENANT_ID=tu_tenant_id_de_azure_ad
MICROSOFT_CLIENT_SECRET=tu_client_secret_de_azure_ad
REDIRECT_URI=http://localhost:5000/auth/callback
3. Instalar dependencias y arrancar
Bash
# Instalar dependencias
npm install

# Levantar entorno de desarrollo
npm run dev
⚠️ Nota Técnica para el Administrador de SharePoint
Para que el guardado real funcione correctamente en producción, asegúrate de crear una columna personalizada en tu biblioteca o lista de SharePoint con las siguientes especificaciones:

Nombre de la columna: AI Tags (SharePoint creará el nombre interno estático como AITags).

Tipo de columna: Multiple lines of text (Texto de varias líneas).

Configuración de texto: Plain text (Texto sin formato).

🧑‍💻 Sobre el Autor (Marca RIEA)
RIEA representa la intersección estratégica entre Gestión del Talento (HR), Operaciones de TI e Inteligencia Artificial.

Como profesional de Recursos Humanos enfocado en la adopción de soluciones de TI y desarrollo de software, creé este producto para demostrar cómo la tecnología moderna e híbrida puede romper las barreras administrativas tradicionales, aportando un verdadero valor de arquitectura, optimización de costos (FinOps) y automatización de procesos a las organizaciones de nivel Enterprise.

Desarrollado con pasión por RIEA | Innovación en Gestión de Datos y Soluciones de IA.

## Development Notice

This project was designed, architected, reviewed and validated by a human.

AI was used as a development assistant for code generation, documentation and refactoring.

RIEA | Asanaé ♠ | by PAMB

Applied Intellectual and Ethical Responsibility
