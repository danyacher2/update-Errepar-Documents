# 📘 Metadata Manager Process - Documentación Completa

**Versión:** 1.0  
**Fecha:** Diciembre 2024  
**Framework:** .NET 8.0  
**Lenguaje:** C# 12.0

---

## 📑 Tabla de Contenidos

1. [Resumen Ejecutivo](#resumen-ejecutivo)
2. [Arquitectura del Sistema](#arquitectura-del-sistema)
3. [Componentes Principales](#componentes-principales)
4. [Flujo de Trabajo](#flujo-de-trabajo)
5. [Sistema de Logs](#sistema-de-logs)
6. [Integración con Sistemas Externos](#integración-con-sistemas-externos)
7. [Configuración y Seguridad](#configuración-y-seguridad)
8. [Troubleshooting](#troubleshooting)
9. [Diagramas de Flujo](#diagramas-de-flujo)

---

## 🎯 Resumen Ejecutivo

**Metadata Manager Process** es un sistema automatizado que gestiona la sincronización de metadatos entre SharePoint y sistemas de búsqueda externos (Solr y Milvus). El proceso:

- ✅ Obtiene items pendientes de procesamiento desde SharePoint
- ✅ Aplica cambios de metadatos a documentos
- ✅ Sincroniza los cambios a Solr (motor de búsqueda)
- ✅ Sincroniza los cambios a Milvus (base de datos vectorial para IA)
- ✅ Genera logs detallados en archivos locales y SharePoint
- ✅ Respeta ventanas de ejecución programadas (Schedule Gates)
- ✅ Maneja reintentos automáticos en caso de fallas

### Características Clave

| Característica | Descripción |
|---------------|-------------|
| **Autenticación** | OAuth con certificados (Azure AD) |
| **Procesamiento** | Batch de múltiples items |
| **Reintentos** | 3 intentos con delays configurables |
| **Logging** | Dual (local + SharePoint) |
| **Scheduling** | Ventanas horarias configurables |
| **Ambientes** | UAT y Producción |

---

## 🏗️ Arquitectura del Sistema

### Diagrama de Componentes

```
┌─────────────────────────────────────────────────────────────┐
│                    METADATA MANAGER                          │
│                      (Program.cs)                            │
└──────────────┬──────────────────────────────────────────────┘
               │
    ┌──────────┴──────────┬──────────────┬────────────────┐
    │                     │              │                │
    ▼                     ▼              ▼                ▼
┌─────────┐      ┌──────────────┐  ┌─────────┐    ┌──────────┐
│SharePoint│      │TokenProvider │  │Schedule │    │Logs      │
│Manager  │      │(Auth)        │  │Gate     │    │Manager   │
└────┬────┘      └──────────────┘  └─────────┘    └────┬─────┘
     │                                                   │
     │                                                   │
┌────┴─────────────────────────────────────────────────┴────┐
│              SharePoint Online                              │
│   Lista: "Metadata Manager"                                │
│   - Estado: Pendiente / En Pausa / Procesando / Finalizado │
│   - Cambios: JSON con metadatos a actualizar               │
│   - Activos: Array de GUIDs de documentos                  │
└──────────────────────┬──────────────────────────────────────┘
                       │
        ┌──────────────┼──────────────┐
        │              │              │
        ▼              ▼              ▼
┌──────────────┐  ┌─────────┐  ┌──────────┐
│ProcesarCambios│  │Solr     │  │Milvus    │
│               │  │Manager  │  │Manager   │
└───────────────┘  └────┬────┘  └────┬─────┘
                        │            │
                        ▼            ▼
                  ┌──────────┐  ┌──────────┐
                  │Solr      │  │Milvus    │
                  │(Búsqueda)│  │(Vector DB│
                  └──────────┘  └──────────┘
```

### Stack Tecnológico

```yaml
Backend:
  - Runtime: .NET 8.0
  - Language: C# 12.0
  - Libraries:
    - Microsoft.SharePoint.Client (CSOM)
    - System.Net.Http
    - System.Text.Json
    - Newtonsoft.Json

Infraestructura:
  - SharePoint Online
  - Apache Solr (UAT/Producción)
  - Milvus Vector Database
  - Azure AD (Autenticación)

Seguridad:
  - OAuth 2.0 con Certificados
  - TLS 1.2/1.3
  - VPN corporativa (opcional)
```

---

## 🔧 Componentes Principales

### 1. **Program.cs** - Orquestador Principal

**Responsabilidades:**
- Inicialización de servicios
- Bucle principal de procesamiento
- Coordinación entre componentes
- Manejo de errores globales

**Flujo de Ejecución:**

```csharp
1. Autenticación (TokenProvider)
   ↓
2. Configuración HttpClient (SSL, DNS, Timeouts)
   ↓
3. Obtener items pendientes (SharePointManager)
   ↓
4. Cargar Schedule (ScheduleGate)
   ↓
5. Por cada Item:
   5.1. Verificar Schedule
   5.2. Buscar activo en SharePoint
   5.3. Aplicar cambios (ProcesarCambios)
   5.4. Sincronizar a Solr (SolrManager)
   5.5. Sincronizar a Milvus (MilvusManager)
   5.6. Guardar logs
   ↓
6. Finalización y sincronización de logs
```

**Configuración HttpClient (Solución SSL/VPN):**

```csharp
var socketHandler = new SocketsHttpHandler
{
    SslOptions = new SslClientAuthenticationOptions
    {
        RemoteCertificateValidationCallback = (sender, cert, chain, errors) => true,
        EnabledSslProtocols = SslProtocols.Tls12 | SslProtocols.Tls13,
        CertificateRevocationCheckMode = X509RevocationMode.NoCheck
    },
    UseProxy = false,                              // Evita conflictos con VPN
    PooledConnectionLifetime = TimeSpan.FromMinutes(5),
    MaxConnectionsPerServer = 20,
    ConnectTimeout = TimeSpan.FromSeconds(30)
};
```

**Razones:**
- ✅ Bypass SSL para certificados UAT auto-firmados
- ✅ Deshabilita proxy que interfiere con VPN corporativa
- ✅ Timeouts largos para conexiones lentas
- ✅ Pool de conexiones optimizado

---

### 2. **SharePointManager** - Gestor de SharePoint

**Métodos Principales:**

| Método | Descripción |
|--------|-------------|
| `GetPendingOrPausedAsync()` | Obtiene items con estado "Pendiente" o "En Pausa" |
| `UpdateEstadoProcesoAsync()` | Actualiza el estado del proceso |
| `UpdateCantActivosAsync()` | Actualiza contador de activos procesados |

**Ejemplo de Uso:**

```csharp
var items = await sp.GetPendingOrPausedAsync(cts.Token);
// Retorna: List<ItemCambios>
// - Id: ID del item en SharePoint
// - Titulo: Descripción del cambio
// - EstadoProceso: "Pendiente" / "En Pausa" / "Procesando" / "Finalizado"
// - Cambios: JsonDocument con los metadatos a cambiar
// - Activos: List<string> con GUIDs de documentos
// - Scope: "Solr" / "Milvus" / "Milvus - IA"
```

---

### 3. **SolrManager** - Gestor de Apache Solr

**Propósito:** Sincronizar cambios de metadatos al índice de búsqueda Solr.

**Métodos Principales:**

#### `AtomicUpdateSolrMultiple()`

```csharp
public async Task<(bool ok, string msg)> AtomicUpdateSolrMultiple(
    string guid,              // GUID del documento
    Dictionary<string, object> campos,  // Campos a actualizar
    int idElemento            // ID para logging
)
```

**Características:**
- ✅ **Actualización atómica** (no reemplaza todo el documento)
- ✅ **3 reintentos** con delay de 5 segundos
- ✅ **Logging detallado** de intentos
- ✅ **Formateo automático** de tipos (bool, DateTime, arrays)

**Formato JSON Enviado:**

```json
[
  {
    "id": "{guid}",
    "eolShpTema": { "set": "Impuestos" },
    "eolShpFecha": { "set": "2024-01-15T10:30:00Z" },
    "eolShpOrganismo": { "set": ["AFIP", "DGI"] }
  }
]
```

#### `GetItemFormatSync()`

**Propósito:** Convierte un `ListItem` de SharePoint a formato compatible con Solr/Milvus.

**Procesamiento de Campos:**

| Tipo de Campo SharePoint | Transformación |
|-------------------------|----------------|
| **Metadatos Administrados** (Single) | `_taxonomyValue.Label` |
| **Metadatos Administrados** (Multi) | `Array<string>` de labels |
| **Búsqueda** (Lookup) | `_lookupValue.LookupValue` |
| **Búsqueda Multi** | `Array<string>` de valores |
| **Número** | `Convert.ToInt32()` |
| **Fecha y Hora** | `DateTime` |
| **Texto** | `string` |

**Casos Especiales:**

```csharp
// 1. Organismo: Extrae último segmento después de ":"
"País: Argentina: Buenos Aires: CABA" → "CABA"

// 2. Índice de Contenidos: Obtiene full path de GUIDs
await getFullPathIds(termGuid, "Índice de Contenidos")

// 3. Legislación: Campo especial "eolShpFechaBusqueda"
if (contentType == "Legislacion" && field == "eolShpFecha")
    col["eolShpFechaBusqueda"] = item["eolShpFecha"];
```

---

### 4. **MilvusManager** - Gestor de Milvus

**Propósito:** Sincronizar documentos a Milvus (base de datos vectorial para búsqueda semántica con IA).

**Método Principal:**

```csharp
public static async Task<(bool ok, string msg)> EnviarAMilvus(
    HttpClient client,
    string json,              // Item formateado completo
    int itemId,
    bool useIA                // true = "Milvus - IA", false = "Milvus"
)
```

**Diferencias entre Modos:**

| Modo | URL Endpoint | Procesamiento |
|------|-------------|---------------|
| **Milvus** | `/api/milvus/update` | Vectorización básica |
| **Milvus - IA** | `/api/milvus/ai/update` | Vectorización avanzada con LLM |

---

### 5. **ProcesarCambios** - Aplicador de Cambios

**Propósito:** Aplica los cambios de metadatos a los items de SharePoint antes de sincronizar.

**Estructura de Cambios (JSON):**

```json
[
  {
    "campo": "eolShpTema",
    "tipoValor": "string",
    "valor": "Derecho Laboral"
  },
  {
    "campo": "eolShpOrganismo",
    "tipoValor": "TaxonomyMulti",
    "valor": "AFIP;DGI;ANSES"
  }
]
```

**Tipos de Valores Soportados:**

- `string` - Texto plano
- `int` - Números enteros
- `DateTime` - Fechas
- `Taxonomy` - Campo de taxonomía único
- `TaxonomyMulti` - Campo de taxonomía múltiple
- `Lookup` - Campo de búsqueda
- `LookupMulti` - Campo de búsqueda múltiple

---

### 6. **LogsManager** - Sistema de Logs

Ver sección completa: [Sistema de Logs](#sistema-de-logs)

---

### 7. **ScheduleGate** - Control de Horarios

**Propósito:** Permite configurar ventanas horarias en las que el proceso puede ejecutarse.

**Configuración en SharePoint:**

Lista: `"Manage Metadata Schedule"`

| Campo | Tipo | Ejemplo |
|-------|------|---------|
| `DayOfWeek` | Number | 1 (Lunes) - 7 (Domingo) |
| `StartTime` | Text | "09:00" |
| `EndTime` | Text | "18:00" |
| `Type` | Choice | "Allow" / "Deny" |

**Lógica:**

```
1. Si NO hay reglas → Permitir siempre
2. Si hora actual está en ALLOW → Permitir
3. Si hora actual está en DENY → Pausar
4. Si no está en ninguna → Depende de la política
```

**Ejemplo:**

```
Allow: Lunes-Viernes 09:00-18:00
Deny:  Sábado-Domingo todo el día

Resultado:
- Lunes 10:00 → ✅ Procesa
- Lunes 20:00 → ⏸️ Pausa hasta el día siguiente 09:00
- Sábado 12:00 → ⏸️ Pausa hasta Lunes 09:00
```

**Método WaitUntilAllowedAsync():**

Hace polling cada 1 minuto hasta que llegue una ventana permitida.

---

## 📋 Sistema de Logs

### Arquitectura de Logs

```
                    ┌──────────────────┐
                    │   LogsManager    │
                    └────────┬─────────┘
                             │
                   ┌─────────┴─────────┐
                   │                   │
                   ▼                   ▼
          ┌─────────────────┐   ┌──────────────┐
          │  ARCHIVOS LOCAL │   │  SHAREPOINT  │
          │  /Logs/Item_X/  │   │  Adjuntos    │
          └─────────────────┘   └──────────────┘
```

### Tipos de Logs

#### 1. **LogEjecucion** - Eventos del Proceso

**Archivo:** `/Logs/Item_{id}/LogEjecucion.json`

**Estructura:**

```json
[
  {
    "ItemId": 110,
    "Fecha": "2024-12-10T14:30:00",
    "Mensaje": "Inicio del proceso Metadata Manager",
    "Estado": "Inicio"
  },
  {
    "ItemId": 110,
    "Fecha": "2024-12-10T14:30:15",
    "Mensaje": "Items recuperados desde SharePoint: 5",
    "Estado": "OK"
  }
]
```

**Cuándo se Usa:**

```csharp
// Inicio del proceso
await logsManager.LogEjecucionAsync(0, "Inicio del proceso", "Inicio");

// Proceso exitoso
await logsManager.LogEjecucionAsync(itemId, "Activo modificado en Solr", "OK");

// Cambio de estado
await logsManager.LogEjecucionAsync(itemId, "Proceso en Pausa", "En Pausa");
```

---

#### 2. **LogError** - Errores

**Archivo:** `/Logs/Item_{id}/LogError.json`

**Estructura:**

```json
[
  {
    "ItemId": "{guid-del-activo}",
    "Fecha": "2024-12-10T14:35:00",
    "Mensaje": "Error al actualizar en Solr: 500 Internal Server Error",
    "Estado": "Error Solr"
  }
]
```

**Cuándo se Usa:**

```csharp
// Error en Solr
await logsManager.LogErrorAsync(
    itemId, 
    activo, 
    "Error al actualizar en Solr: {msg}", 
    "Error Solr"
);

// Item no encontrado
await logsManager.LogErrorAsync(
    itemId, 
    activo, 
    "Item con guid {activo} no se encontro", 
    "No Encontrado"
);
```

---

#### 3. **ActivosProcesados** - Detalle de Items

**Archivo:** `/Logs/Item_{id}/ActivosProcesados.json`

**Estructura:**

```json
[
  {
    "ItemCambiosId": 110,
    "ItemListId": 2456,
    "TimeStamp": 2456,
    "UrlItem": "/sites/ErreparDesarrollo/Legislacion/ley-27430.aspx",
    "LibraryName": "Legislación",
    "Activo": "a1b2c3d4-e5f6-7890-1234-567890abcdef",
    "Procesado": true,
    "Fecha": "2024-12-10T14:35:30"
  }
]
```

**Cuándo se Usa:**

```csharp
var logActivo = new ItemLogActivosProcesados
{
    ItemCambiosId = it.Id,
    ItemListId = item.Id,
    LibraryName = item["Title"]?.ToString(),
    UrlItem = item["FileRef"]?.ToString(),
    Activo = activo,
    Procesado = true,
    Fecha = DateTime.Now
};

await logsManager.SaveLogActivosProcesados(itemId, new[] { logActivo });
```

---

#### 4. **Snapshots de Items**

**Archivos:**
- `/searchLogs/ListItems/SearchOriginal_{guid}.json` - Antes de procesar
- `/searchLogs/ListItems/Search_{guid}.json` - Después de procesar

**Propósito:** Auditoría de cambios (antes/después)

```csharp
// Antes de aplicar cambios
await logsManager.SaveItemOriginalJsonAsync(activo, item.FieldValues);

// Después de aplicar cambios
await logsManager.SaveItemProcessedJsonAsync(activo, item.FieldValues);
```

---

### Sincronización de Logs a SharePoint

**Método:** `SyncLogs(int itemId, CancellationToken ct)`

**Proceso:**

```
1. Lee los 3 archivos JSON del disco
   - LogEjecucion.json
   - ActivosProcesados.json
   - LogError.json

2. Por cada archivo:
   a. GET /_api/web/lists/.../AttachmentFiles
      → Verifica si el adjunto ya existe
   
   b. Si existe:
      DELETE /_api/web/lists/.../AttachmentFiles('{fileName}')
      → Borra el adjunto antiguo
   
   c. POST /_api/web/lists/.../AttachmentFiles/add(FileName='{fileName}')
      Content-Type: application/json
      Body: [contenido del archivo]
      → Sube el nuevo adjunto

3. Reintentos: 3 intentos con delay de 1-2 segundos
```

**Resultado en SharePoint:**

El item en la lista "Metadata Manager" tendrá 3 adjuntos:
- 📄 LogEjecucion.json
- 📄 ActivosProcesados.json
- 📄 LogError.json

---

### Estructura de Carpetas Completa

```
/Logs
  /Item_1
    ├─ LogEjecucion.json          [✅ Se sube a SharePoint]
    ├─ LogError.json               [✅ Se sube a SharePoint]
    └─ ActivosProcesados.json      [✅ Se sube a SharePoint]
  /Item_2
    ├─ LogEjecucion.json
    ├─ LogError.json
    └─ ActivosProcesados.json

/searchLogs
  /ListItems
    ├─ SearchOriginal_{guid1}.json  [❌ Solo local]
    ├─ Search_{guid1}.json          [❌ Solo local]
    ├─ SearchOriginal_{guid2}.json
    └─ Search_{guid2}.json
```

---

## 🔄 Flujo de Trabajo Completo

### Diagrama de Secuencia

```
Usuario          Program.cs      SharePointManager    SolrManager    MilvusManager    LogsManager
   │                 │                   │                │               │               │
   │─────Ejecuta─────>│                  │                │               │               │
   │                 │                   │                │               │               │
   │                 │──GetPending──────>│                │               │               │
   │                 │<─────Items────────│                │               │               │
   │                 │                   │                │               │               │
   │                 │─────────────────────────Log("Inicio")──────────────>│
   │                 │                   │                │               │<──SaveLocal───│
   │                 │                   │                │               │               │
   │                 ├──┐                │                │               │               │
   │                 │  │ For cada item  │                │               │               │
   │                 │<─┘                │                │               │               │
   │                 │                   │                │               │               │
   │                 │──FindItem────────>│                │               │               │
   │                 │<─────Item─────────│                │               │               │
   │                 │                   │                │               │               │
   │                 │──ProcesarCambios─>│                │               │               │
   │                 │<─────OK───────────│                │               │               │
   │                 │                   │                │               │               │
   │                 │──AtomicUpdate─────────────────────>│               │               │
   │                 │                   │                │─Intento 1────>│               │
   │                 │                   │                │<────OK────────│               │
   │                 │<──────OK──────────────────────────│               │               │
   │                 │                   │                │               │               │
   │                 │──EnviarMilvus──────────────────────────────────────>│              │
   │                 │                   │                │               │─API Call────>│
   │                 │<──────OK──────────────────────────────────────────│               │
   │                 │                   │                │               │               │
   │                 │────────────────────────────Log("OK Solr")─────────────────────────>│
   │                 │────────────────────────────Log("OK Milvus")───────────────────────>│
   │                 │                   │                │               │               │
   │                 │──SyncLogs─────────────────────────────────────────────────────────>│
   │                 │                   │                │               │─Upload ShareP─>│
   │                 │                   │                │               │               │
   │<────Finalizado──│                   │                │               │               │
```

### Flujo Detallado Paso a Paso

#### **Fase 1: Inicialización**

```
1. TokenProvider.GetSharePointTokenWithCertificateThumbprintAsync()
   → Obtiene OAuth token usando certificado de Azure AD
   
2. new HttpClient(socketHandler)
   → Configura HttpClient con bypass SSL y configuración VPN
   
3. new ClientContext(siteUrl)
   → Inicializa contexto de SharePoint CSOM
   
4. new SolrManager(http, siteUrl, "UAT")
   → Inicializa gestor de Solr
   
5. new LogsManager(http, siteUrl, "Metadata Manager")
   → Inicializa gestor de logs
   
6. SharePointManager.GetPendingOrPausedAsync()
   → Obtiene lista de items a procesar
   
7. ScheduleGate.LoadWindows()
   → Carga ventanas horarias permitidas
```

#### **Fase 2: Procesamiento (Por cada Item)**

```
FOR cada ItemCambios en items:
  
  1. ManejarPausaPorSchedule()
     → Verifica si la hora actual está permitida
     → Si NO → Pausa hasta ventana permitida
     
  2. FOR cada activo (GUID) en item.Activos:
     
     2.1. FindItem(listas, guid)
          → Busca el item en las bibliotecas de SharePoint
          → Intenta primero en listaProbable (cache)
          → Si no existe, busca en todas las bibliotecas
          
     2.2. Filtros:
          → Skip si es carpeta
          → Skip si es imagen (.png, .jpg, .jpeg, .gif)
          
     2.3. SaveItemOriginalJsonAsync()
          → Guarda snapshot "antes de cambios"
          
     2.4. ProcesarCambios.Procesar()
          → Aplica los cambios de metadatos
          → item.SystemUpdate() → Guarda en SharePoint
          
     2.5. SI "Solr" en it.Scope:
          ┌─────────────────────────────────┐
          │ ProcesarEnSolr()                │
          │  → ObtenerItemFormateadoYCampos │
          │  → AtomicUpdateSolrMultiple     │
          │     → 3 reintentos con delay    │
          └─────────────────────────────────┘
          
     2.6. SI "Milvus" en it.Scope:
          ┌─────────────────────────────────┐
          │ ProcesarEnMilvusAsync()         │
          │  → ObtenerItemFormateadoYCampos │
          │  → EnviarAMilvus(useIA)         │
          └─────────────────────────────────┘
          
     2.7. SaveItemProcessedJsonAsync()
          → Guarda snapshot "después de cambios"
          
     2.8. SaveLogActivosProcesados()
          → Registra el activo procesado
          
  3. UpdateEstadoProcesoAsync("Finalizado")
     → Marca el item como completado
     
  4. SyncLogs()
     → Sube los 3 archivos JSON a SharePoint
```

#### **Fase 3: Finalización**

```
1. LogEjecucionAsync(0, "Proceso completado", "Finalizado")
   → Registra finalización exitosa
   
2. Console.WriteLine("Presiona una tecla...")
   → Espera confirmación del usuario
```

---

## 🔗 Integración con Sistemas Externos

### SharePoint Online

**API Utilizada:** CSOM (Client-Side Object Model)

**Autenticación:**

```csharp
// 1. Obtener token OAuth con certificado
var authToken = await TokenProvider.GetSharePointTokenWithCertificateThumbprintAsync(
    tenantId, 
    clientId, 
    certificateThumbprint, 
    siteUrl
);

// 2. Inyectar token en cada request CSOM
context.ExecutingWebRequest += (sender, e) =>
{
    e.WebRequestExecutor.RequestHeaders["Authorization"] = "Bearer " + authToken;
};
```

**Operaciones Principales:**

```csharp
// Buscar item por GUID
var query = new CamlQuery
{
    ViewXml = $@"
        <View Scope='RecursiveAll'>
            <Query>
                <Where>
                    <Eq>
                        <FieldRef Name='GUID' />
                        <Value Type='Guid'>{uniqueId}</Value>
                    </Eq>
                </Where>
            </Query>
            <RowLimit>1</RowLimit>
        </View>"
};

var items = list.GetItems(query);
context.Load(items);
context.ExecuteQuery();

// Actualizar metadatos
item["eolShpTema"] = "Nuevo Valor";
item.SystemUpdate(); // No actualiza "Modified" / "ModifiedBy"
context.ExecuteQuery();
```

---

### Apache Solr

**URL:** `https://solr.uat.errepar.com/solr/prodActivos02/update?commit=true`

**Autenticación:**

```csharp
string credentials = ambiente == "PROD" 
    ? "admin:6s2HUXFb8la" 
    : "uat-solr-acess:SXtuCde9&JW%pkAK";

string authHeader = "Basic " + Base64Encode(credentials);
httpRequestMessage.Headers.Add("Authorization", authHeader);
```

**Atomic Update (Partial Update):**

```http
POST /solr/prodActivos02/update?commit=true
Content-Type: application/json
Authorization: Basic {base64}

[
  {
    "id": "a1b2c3d4-e5f6-7890-1234-567890abcdef",
    "eolShpTema": { "set": "Impuestos" },
    "eolShpFecha": { "set": "2024-01-15T10:30:00Z" },
    "eolShpOrganismo": { "set": ["AFIP", "DGI"] }
  }
]
```

**Ventajas del Atomic Update:**
- ✅ Solo actualiza campos especificados
- ✅ No requiere re-indexar todo el documento
- ✅ Más rápido y eficiente

---

### Milvus (Vector Database)

**URL Base:** Configurada en `MilvusManager`

**Endpoints:**

```
POST /api/milvus/update          → Milvus estándar
POST /api/milvus/ai/update       → Milvus con IA (vectorización avanzada)
```

**Payload:**

```json
{
  "id": "a1b2c3d4-e5f6-7890-1234-567890abcdef",
  "eolShpTipoContenido": "Legislacion",
  "eolShpTitle": "Ley 27.430 - Reforma Tributaria",
  "eolShpBody": "Contenido completo del documento...",
  "eolShpTema": "Impuestos",
  "eolShpOrganismo": ["AFIP", "DGI"],
  ...
}
```

**Procesamiento:**
1. Milvus recibe el JSON completo
2. Genera embeddings vectoriales del contenido
3. Almacena en base de datos vectorial
4. Permite búsqueda semántica (búsqueda por significado, no solo palabras clave)

---

## 🔐 Configuración y Seguridad

### Variables de Entorno

```bash
# Opcional: Sobrescribe el thumbprint del certificado
CERT_THUMBPRINT=452079A2697BC9646023FAE02876488654BBDB2C
```

### Certificados Azure AD

**Requisitos:**
- Certificado .pfx instalado en el equipo
- Registrado en Azure AD App Registration
- Permisos: `Sites.FullControl.All`

**Ubicación:**
```
CurrentUser\My\{thumbprint}
```

**Validación:**

```powershell
# Listar certificados instalados
Get-ChildItem -Path Cert:\CurrentUser\My

# Verificar permisos
Get-MgApplication -ApplicationId {clientId}
```

---

### Configuración de Red (VPN/SSL)

**Problema Común:** "The SSL connection could not be established" o "The requested name is valid, but no data of the requested type was found"

**Solución Implementada:**

```csharp
var socketHandler = new SocketsHttpHandler
{
    SslOptions = new SslClientAuthenticationOptions
    {
        // Bypass validación SSL (solo UAT)
        RemoteCertificateValidationCallback = (sender, cert, chain, errors) => true,
        
        // Forzar TLS 1.2/1.3
        EnabledSslProtocols = SslProtocols.Tls12 | SslProtocols.Tls13,
        
        // No verificar revocación de certificados (VPN puede bloquear)
        CertificateRevocationCheckMode = X509RevocationMode.NoCheck
    },
    
    // Deshabilitar proxy (evita conflictos con VPN)
    UseProxy = false,
    
    // Pool de conexiones
    PooledConnectionLifetime = TimeSpan.FromMinutes(5),
    MaxConnectionsPerServer = 20,
    ConnectTimeout = TimeSpan.FromSeconds(30)
};
```

**Para Producción:**

⚠️ **NO usar bypass SSL en producción**. Configurar certificados válidos.

---

### Permisos SharePoint

**App Permissions (Azure AD):**

```
Sites.FullControl.All (Application)
```

**Permisos de Lista:**

```
Lista: "Metadata Manager"
- Read: ✅
- Write: ✅ (actualizar Estado, CantActivos)
- Attach: ✅ (adjuntar logs)

Bibliotecas de Documentos:
- Read: ✅
- Write: ✅ (actualizar metadatos)
```

---

## 🐛 Troubleshooting

### Errores SSL/TLS

#### **Error 1: "The SSL connection could not be established"**

**Causa:** Certificado SSL del servidor Solr no es válido (auto-firmado en UAT)

**Solución:**

```csharp
RemoteCertificateValidationCallback = (sender, cert, chain, errors) => true
```

**Verificar:**

```powershell
Test-NetConnection -ComputerName solr.uat.errepar.com -Port 443
```

---

#### **Error 2: "The requested name is valid, but no data of the requested type was found"**

**Causa:** Problema de resolución DNS (VPN solo configura IPv4)

**Solución:**

```csharp
// Configuración global DNS
ServicePointManager.DnsRefreshTimeout = 0;
ServicePointManager.EnableDnsRoundRobin = true;
```

**Verificar:**

```powershell
nslookup solr.uat.errepar.com
# Debe retornar dirección IPv4
```

---

#### **Error 3: "Received an unexpected EOF or 0 bytes from the transport stream"**

**Causa:** Timeout SSL o interrupción de firewall/proxy

**Solución:**

```csharp
ConnectTimeout = TimeSpan.FromSeconds(30)  // Aumentar timeout
```

---

### Errores de Autenticación

#### **Error 4: "401 Unauthorized" en SharePoint**

**Causa:** Token OAuth expirado o certificado inválido

**Diagnóstico:**

```csharp
// Agregar logging
Console.WriteLine($"Token: {authToken.Substring(0, 20)}...");
Console.WriteLine($"Cert Thumbprint: {certificateThumbprint}");
```

**Solución:**

1. Verificar que el certificado esté instalado:

```powershell
Get-ChildItem -Path Cert:\CurrentUser\My | 
    Where-Object { $_.Thumbprint -eq "452079A2697BC9646023FAE02876488654BBDB2C" }
```

2. Verificar permisos en Azure AD:

```
Azure Portal → App Registrations → {App} → API Permissions
- Sites.FullControl.All (Application)
```

---

### Errores de Procesamiento

#### **Error 5: "Item con guid {X} no se encontro"**

**Causa:** El GUID del activo no existe en ninguna biblioteca o fue eliminado

**Diagnóstico:**

```csharp
// Agregar logging detallado en FindItem()
foreach (var lista in listas)
{
    Console.WriteLine($"Buscando en lista: {lista.Title}");
    var item = GetItemByUniqueId(ctx, lista, guid);
    if (item != null) 
    {
        Console.WriteLine($"✅ Encontrado en {lista.Title}");
        return item;
    }
}
Console.WriteLine($"❌ No encontrado en ninguna lista");
```

**Solución:**

1. Verificar que el GUID sea correcto en la lista "Metadata Manager"
2. Verificar que el item no haya sido eliminado
3. Verificar que el item no esté en la papelera de reciclaje

---

#### **Error 6: "⚠️ El campo '{campo}' no existe en ItemFormateado"**

**Causa:** El campo especificado en "Cambios" no existe en el documento de SharePoint

**Diagnóstico:**

```csharp
// Listar campos disponibles
foreach (Field field in item.ParentList.Fields)
{
    if (field.InternalName.StartsWith("eolShp"))
        Console.WriteLine($"Campo: {field.InternalName}");
}
```

**Solución:**

1. Verificar el InternalName del campo en SharePoint
2. Actualizar el JSON de cambios con el nombre correcto
3. Verificar que el campo exista en el ContentType del item

---

### Errores de Solr

#### **Error 7: "SOLR - ERROR en activo {id} | 500 Internal Server Error"**

**Causa:** Error interno de Solr (schema, parsing, etc.)

**Diagnóstico:**

```csharp
// Ver response body completo
var responseBody = await response.Content.ReadAsStringAsync();
Console.WriteLine($"Response body: {responseBody}");
```

**Solución:**

1. Verificar el schema de Solr (`schema.xml`)
2. Verificar que los campos existan en Solr
3. Verificar formato del JSON (tipos de datos)

---

#### **Error 8: "Falló después de 3 intentos. Último error: {msg}"**

**Causa:** Solr no responde o hay problemas de red persistentes

**Solución:**

1. Aumentar número de reintentos:

```csharp
const int maxIntentos = 5;  // En lugar de 3
```

2. Aumentar delay entre reintentos:

```csharp
const int delayEntreIntentos = 10000;  // 10 segundos
```

3. Verificar que Solr esté corriendo:

```bash
curl https://solr.uat.errepar.com/solr/admin/ping
```

---

### Errores de Milvus

#### **Error 9: "Error al actualizar en Milvus: {msg}"**

**Diagnóstico:**

```csharp
// Agregar logging del JSON enviado
Console.WriteLine($"JSON enviado a Milvus:\n{json}");
```

**Solución:**

1. Verificar que el servicio Milvus esté corriendo
2. Verificar formato del JSON
3. Verificar que el endpoint sea correcto

---

### Errores de Schedule

#### **Error 10: Proceso se queda en "En Pausa" indefinidamente**

**Causa:** No hay ventanas ALLOW configuradas o las reglas están mal definidas

**Diagnóstico:**

```csharp
Console.WriteLine($"Allow rules: {allow.Count}");
Console.WriteLine($"Deny rules: {deny.Count}");

foreach (var rule in allow)
{
    Console.WriteLine($"Allow: {rule.DayOfWeek} {rule.StartTime}-{rule.EndTime}");
}
```

**Solución:**

1. Verificar que existan reglas ALLOW en la lista "Manage Metadata Schedule"
2. Verificar que la hora actual esté en una ventana permitida
3. Para debugging, comentar temporalmente la verificación:

```csharp
// Comentar temporalmente para testing
// EstadoProceso = await ManejarPausaPorSchedule(...);
```

---

## 📊 Diagramas de Flujo

### Flujo General del Sistema

```
┌───────────────────────────────────────────────────────────────┐
│                         INICIO                                 │
│                  (Program.cs Main)                             │
└───────────────────────────┬───────────────────────────────────┘
                            │
                            ▼
┌───────────────────────────────────────────────────────────────┐
│ 1. AUTENTICACIÓN                                               │
│    - Obtener OAuth Token (Certificado Azure AD)               │
│    - Configurar HttpClient (SSL, VPN, Timeouts)               │
│    - Inicializar ClientContext SharePoint                     │
└───────────────────────────┬───────────────────────────────────┘
                            │
                            ▼
┌───────────────────────────────────────────────────────────────┐
│ 2. INICIALIZACIÓN DE SERVICIOS                                │
│    - SolrManager                                               │
│    - MilvusManager                                             │
│    - LogsManager                                               │
│    - SharePointManager                                         │
│    - ProcesarCambios                                           │
└───────────────────────────┬───────────────────────────────────┘
                            │
                            ▼
┌───────────────────────────────────────────────────────────────┐
│ 3. OBTENER DATOS                                               │
│    - GetPendingOrPausedAsync()                                 │
│      → List<ItemCambios> con Estado = "Pendiente" | "En Pausa"│
│    - LoadWindows(ScheduleGate)                                 │
│      → Allow/Deny rules                                        │
└───────────────────────────┬───────────────────────────────────┘
                            │
                            ▼
                    ┌───────────────┐
                    │  FOR EACH     │
                    │  ItemCambios  │
                    └───────┬───────┘
                            │
                            ▼
┌───────────────────────────────────────────────────────────────┐
│ 4. VERIFICAR SCHEDULE                                          │
│    ┌──────────────────────────────────────────────────┐      │
│    │ ¿Hora actual en ventana ALLOW?                   │      │
│    └───────┬──────────────────────────┬───────────────┘      │
│            │ NO                        │ SI                    │
│            ▼                           ▼                       │
│    ┌────────────────┐        ┌─────────────────┐            │
│    │ Pausa Proceso  │        │ Continuar       │            │
│    │ UpdateEstado   │        │                 │            │
│    │ "En Pausa"     │        │                 │            │
│    │                │        │                 │            │
│    │ Wait 1 min     │        │                 │            │
│    └────────┬───────┘        └────────┬────────┘            │
│             │                         │                       │
│             └──────────┬──────────────┘                       │
│                        │                                       │
└────────────────────────┼───────────────────────────────────────┘
                         │
                         ▼
                ┌────────────────┐
                │  FOR EACH      │
                │  Activo (GUID) │
                └────────┬───────┘
                         │
                         ▼
┌───────────────────────────────────────────────────────────────┐
│ 5. BUSCAR ITEM EN SHAREPOINT                                   │
│    FindItem(listas, guid)                                      │
│    ┌─────────────────────────────────────────────────┐       │
│    │ ¿Item encontrado?                               │       │
│    └──────┬──────────────────────┬───────────────────┘       │
│           │ NO                    │ SI                         │
│           ▼                       ▼                            │
│    ┌────────────┐        ┌──────────────────┐               │
│    │ Log Error  │        │ Continue          │               │
│    │ Skip       │        │                   │               │
│    └────────────┘        └──────┬────────────┘               │
│                                  │                             │
└──────────────────────────────────┼─────────────────────────────┘
                                   │
                                   ▼
┌───────────────────────────────────────────────────────────────┐
│ 6. FILTROS                                                     │
│    ┌────────────────────────────────────────────┐            │
│    │ ¿Es carpeta?           → SI → Skip         │            │
│    │ ¿Es imagen (.png/jpg)? → SI → Skip         │            │
│    └────────────────┬───────────────────────────┘            │
│                     │ NO (continuar)                          │
└─────────────────────┼───────────────────────────────────────┘
                      │
                      ▼
┌───────────────────────────────────────────────────────────────┐
│ 7. GUARDAR SNAPSHOT ORIGINAL                                   │
│    SaveItemOriginalJsonAsync()                                 │
│    → /searchLogs/ListItems/SearchOriginal_{guid}.json         │
└───────────────────────┬───────────────────────────────────────┘
                        │
                        ▼
┌───────────────────────────────────────────────────────────────┐
│ 8. APLICAR CAMBIOS EN SHAREPOINT                               │
│    ProcesarCambios.Procesar(item, cambios)                     │
│    - Actualiza metadatos según JSON                            │
│    - item.SystemUpdate()                                        │
│    - context.ExecuteQuery()                                     │
└───────────────────────┬───────────────────────────────────────┘
                        │
                        ▼
┌───────────────────────────────────────────────────────────────┐
│ 9. SINCRONIZAR A SISTEMAS EXTERNOS                             │
│                                                                 │
│    ┌──────────────────────────────────────────┐              │
│    │ ¿Scope contiene "Solr"?                 │              │
│    └────┬─────────────────────┬───────────────┘              │
│         │ SI                   │ NO                            │
│         ▼                      │                               │
│    ┌─────────────────────┐   │                               │
│    │ ProcesarEnSolr()     │   │                               │
│    │ - GetItemFormat      │   │                               │
│    │ - AtomicUpdate       │   │                               │
│    │ - 3 Reintentos       │   │                               │
│    └─────────┬────────────┘   │                               │
│              │                 │                               │
│              ▼                 │                               │
│    ┌─────────────────────┐   │                               │
│    │ ¿Éxito?              │   │                               │
│    └───┬────────────┬─────┘   │                               │
│        │ SI         │ NO       │                               │
│        ▼            ▼          │                               │
│    ┌──────┐   ┌─────────┐   │                               │
│    │ Log  │   │ LogError│   │                               │
│    │ OK   │   │         │   │                               │
│    └──────┘   └─────────┘   │                               │
│                               │                               │
│    ┌──────────────────────────▼───────────────┐             │
│    │ ¿Scope contiene "Milvus"?                │             │
│    └────┬─────────────────────┬────────────────┘             │
│         │ SI                   │ NO                            │
│         ▼                      │                               │
│    ┌─────────────────────┐   │                               │
│    │ ProcesarEnMilvus()   │   │                               │
│    │ - GetItemFormat      │   │                               │
│    │ - EnviarAMilvus      │   │                               │
│    │   (useIA?)           │   │                               │
│    └─────────┬────────────┘   │                               │
│              │                 │                               │
│              ▼                 │                               │
│    ┌─────────────────────┐   │                               │
│    │ ¿Éxito?              │   │                               │
│    └───┬────────────┬─────┘   │                               │
│        │ SI         │ NO       │                               │
│        ▼            ▼          │                               │
│    ┌──────┐   ┌─────────┐   │                               │
│    │ Log  │   │ LogError│   │                               │
│    │ OK   │   │         │   │                               │
│    └──────┘   └─────────┘   │                               │
│                               │                               │
└───────────────────────────────┼───────────────────────────────┘
                                │
                                ▼
┌───────────────────────────────────────────────────────────────┐
│ 10. GUARDAR SNAPSHOT PROCESADO                                 │
│     SaveItemProcessedJsonAsync()                               │
│     → /searchLogs/ListItems/Search_{guid}.json                │
└───────────────────────┬───────────────────────────────────────┘
                        │
                        ▼
┌───────────────────────────────────────────────────────────────┐
│ 11. REGISTRAR ACTIVO PROCESADO                                 │
│     SaveLogActivosProcesados()                                 │
│     → /Logs/Item_{id}/ActivosProcesados.json                  │
└───────────────────────┬───────────────────────────────────────┘
                        │
                        │ (End FOR EACH Activo)
                        │
                        ▼
┌───────────────────────────────────────────────────────────────┐
│ 12. FINALIZAR ITEM                                             │
│     - UpdateEstadoProcesoAsync("Finalizado")                   │
│     - SyncLogs(itemId)                                         │
│       → Sube 3 JSONs a SharePoint como adjuntos               │
└───────────────────────┬───────────────────────────────────────┘
                        │
                        │ (End FOR EACH ItemCambios)
                        │
                        ▼
┌───────────────────────────────────────────────────────────────┐
│ 13. FINALIZACIÓN                                               │
│     - LogEjecucionAsync("Proceso completado", "Finalizado")   │
│     - Console.ReadKey() → Esperar usuario                     │
└───────────────────────┬───────────────────────────────────────┘
                        │
                        ▼
                   ┌─────────┐
                   │   FIN   │
                   └─────────┘
```

---

### Flujo de Reintentos (Solr/Milvus)

```
┌──────────────────────────────────────┐
│ EnviarASolr() / EnviarAMilvus()      │
└───────────────┬──────────────────────┘
                │
                ▼
        ┌───────────────┐
        │ intentoActual │
        │     = 1       │
        └───────┬───────┘
                │
                ▼
        ┌───────────────────────┐
        │ TRY                    │
        │   SendAsync()          │
        └───────┬────────┬───────┘
                │        │
        ┌───────┘        └────────┐
        │ OK                   ERROR│
        ▼                           ▼
┌──────────────┐        ┌──────────────────┐
│ return       │        │ Log Exception    │
│ (true, "200")│        │                  │
└──────────────┘        └────────┬─────────┘
                                 │
                                 ▼
                        ┌─────────────────┐
                        │ intentoActual   │
                        │    < 3?         │
                        └───┬─────────┬───┘
                            │ SI      │ NO
                            ▼         ▼
                    ┌──────────┐  ┌────────────┐
                    │ Delay    │  │ return     │
                    │ 5 seg    │  │ (false, e) │
                    └────┬─────┘  └────────────┘
                         │
                         │
                    ┌────┴─────┐
                    │ intento  │
                    │   += 1   │
                    └────┬─────┘
                         │
                         │ (Loop back)
                         │
                         └──────────────────────┐
                                                │
                                                ▼
                                    ┌──────────────────┐
                                    │ TRY (again)      │
                                    │   SendAsync()    │
                                    └──────────────────┘
```

---

## 📝 Glosario

| Término | Definición |
|---------|------------|
| **CSOM** | Client-Side Object Model - API de SharePoint para .NET |
| **Atomic Update** | Actualización parcial en Solr (solo campos modificados) |
| **Schedule Gate** | Sistema de ventanas horarias permitidas para ejecución |
| **ItemCambios** | Item de la lista "Metadata Manager" con cambios a aplicar |
| **Activo** | Documento en SharePoint identificado por GUID |
| **SystemUpdate** | Método de actualización que NO modifica "Modified"/"ModifiedBy" |
| **Taxonomy** | Sistema de metadatos jerárquicos de SharePoint |
| **Lookup** | Campo de búsqueda que referencia otro item |
| **ContentType** | Tipo de contenido en SharePoint (Legislación, Doctrina, etc.) |
| **Scope** | Campo que indica dónde sincronizar (Solr, Milvus, Milvus - IA) |
| **Vector DB** | Base de datos de embeddings vectoriales para búsqueda semántica |
| **OAuth** | Protocolo de autenticación usado con Azure AD |
| **Thumbprint** | Huella digital (hash) de un certificado |
| **GUID** | Globally Unique Identifier - Identificador único universal |

---

## 📚 Referencias

### Documentación Oficial

- [Microsoft SharePoint CSOM](https://learn.microsoft.com/en-us/sharepoint/dev/sp-add-ins/complete-basic-operations-using-sharepoint-client-library-code)
- [Apache Solr - Atomic Updates](https://solr.apache.org/guide/solr/latest/indexing-guide/partial-document-updates.html)
- [Milvus Documentation](https://milvus.io/docs)
- [.NET 8 HttpClient](https://learn.microsoft.com/en-us/dotnet/api/system.net.http.httpclient)

### Herramientas

- Visual Studio 2022
- Azure Portal
- SharePoint Online Admin Center
- Solr Admin UI
- Postman (testing APIs)

---

## 🔄 Historial de Cambios

### Versión 1.0 - Diciembre 2024

**Mejoras Implementadas:**

1. ✅ **Configuración SSL/VPN robusta**
   - Bypass SSL para UAT
   - Configuración DNS mejorada
   - Timeouts optimizados

2. ✅ **Refactorización de código**
   - Función común `ObtenerItemFormateadoYCampos()`
   - Elimina duplicación de lógica

3. ✅ **Sistema de logs mejorado**
   - Dual (local + SharePoint)
   - Snapshots antes/después
   - Logs de errores detallados

4. ✅ **Reintentos con delays**
   - 3 intentos para Solr/Milvus
   - Delays configurables (5 segundos)
   - Logging de cada intento

5. ✅ **Soporte para Milvus IA**
   - Modo estándar vs IA
   - Vectorización avanzada con LLM

---

## 📧 Contacto y Soporte

**Equipo de Desarrollo:**  
Metadata Manager Team - Errepar

**Repositorio:**  
https://github.com/danyacher/MetadataManagerProcess

**Branch Activo:**  
`Desarrollo`

---

## ⚙️ Configuración de Ambientes

### UAT (User Acceptance Testing)

```yaml
SharePoint:
  URL: https://erreparsa.sharepoint.com/sites/ErreparDesarrollo
  Lista: Metadata Manager
  
Solr:
  URL: https://solr.uat.errepar.com/solr/prodActivos02
  Auth: Basic uat-solr-acess:SXtuCde9&JW%pkAK
  
Azure AD:
  TenantId: 00f26ad1-2073-4746-a79f-c83061db35c0
  ClientId: f679c472-7b0c-45dc-b38c-cca0b662f77a
  Certificado: 452079A2697BC9646023FAE02876488654BBDB2C
```

### Producción

```yaml
SharePoint:
  URL: https://erreparsa.sharepoint.com/sites/ErreparProduccion
  Lista: Metadata Manager
  
Solr:
  URL: https://solr.errepar.com/solr/prodActivos02
  Auth: Basic admin:6s2HUXFb8la
  
Azure AD:
  TenantId: 00f26ad1-2073-4746-a79f-c83061db35c0
  ClientId: 8688eed4-7464-4288-9820-34849fd19296
  Certificado: [Certificado de Producción]
```

---

## 🎓 Guías de Uso

### Agregar un Nuevo Campo a Sincronizar

**1. Actualizar `SolrManager.GetItemFormatSync()`:**

```csharp
// Agregar nuevo campo
if (field.InternalName.Equals("eolShpNuevoCampo"))
{
    col["eolShpNuevoCampo"] = item["eolShpNuevoCampo"];
}
```

**2. Actualizar schema de Solr:**

```xml
<field name="eolShpNuevoCampo" type="string" indexed="true" stored="true"/>
```

**3. Agregar en JSON de cambios:**

```json
{
  "campo": "eolShpNuevoCampo",
  "tipoValor": "string",
  "valor": "Valor ejemplo"
}
```

---

### Agregar una Nueva Biblioteca de SharePoint

**Actualizar `CargarBibliotecasSharePoint()`:**

```csharp
var libraryNames = new List<string>
{
    "Documento",
    "Agenda",
    // ... existentes ...
    "Nueva Biblioteca"  // ← Agregar aquí
};
```

---

### Cambiar Ventanas de Ejecución (Schedule)

**En SharePoint:**

1. Ir a lista: `"Manage Metadata Schedule"`
2. Agregar/Editar items:

| DayOfWeek | StartTime | EndTime | Type  |
|-----------|-----------|---------|-------|
| 1         | 09:00     | 18:00   | Allow |
| 6         | 00:00     | 23:59   | Deny  |

**Resultado:**
- Lunes 09:00-18:00 → Permitido
- Sábado todo el día → Bloqueado

---

## 📈 Métricas y Monitoreo

### KPIs del Proceso

| Métrica | Descripción | Objetivo |
|---------|-------------|----------|
| **Tasa de Éxito** | Items procesados sin errores / Total items | > 95% |
| **Tiempo Promedio** | Tiempo por item procesado | < 10 seg |
| **Reintentos Solr** | % de items que requieren reintentos | < 10% |
| **Disponibilidad** | % uptime del proceso | > 99% |

### Logs a Monitorear

```bash
# Errores críticos
grep "⚠️" /Logs/Item_*/LogError.json

# Items procesados hoy
find /Logs -name "ActivosProcesados.json" -mtime 0

# Reintentos de Solr
grep "Intento 2" /Logs/Item_*/LogEjecucion.json
```

---

## ✅ Checklist de Despliegue

### Pre-Despliegue

- [ ] Certificado instalado en servidor
- [ ] Permisos Azure AD configurados
- [ ] SharePoint lista "Metadata Manager" creada
- [ ] Solr accesible desde servidor
- [ ] Milvus accesible desde servidor
- [ ] VPN configurada (si aplica)

### Despliegue

- [ ] Compilar en modo Release
- [ ] Ejecutar tests unitarios
- [ ] Validar en ambiente UAT
- [ ] Documentar cambios en CHANGELOG.md
- [ ] Merge a branch main
- [ ] Tag de versión (v1.0.0)

### Post-Despliegue

- [ ] Monitorear primeras ejecuciones
- [ ] Verificar logs en SharePoint
- [ ] Validar sincronización a Solr
- [ ] Validar sincronización a Milvus
- [ ] Notificar a stakeholders

---

## 🎯 Roadmap Futuro

### Versión 2.0 (Q1 2025)

- [ ] **Procesamiento paralelo** de activos (Task.WhenAll)
- [ ] **Dashboard de monitoreo** (Grafana/Power BI)
- [ ] **API REST** para ejecución remota
- [ ] **Notificaciones** (email/Teams) en caso de errores
- [ ] **Rollback automático** en caso de fallas críticas

### Versión 3.0 (Q2 2025)

- [ ] **Machine Learning** para detección de anomalías
- [ ] **Auto-tuning** de timeouts y reintentos
- [ ] **Multi-tenancy** (múltiples sitios SharePoint)
- [ ] **Modo dry-run** (simulación sin cambios reales)

---

## 🔒 Consideraciones de Seguridad

### Producción

⚠️ **IMPORTANTE:**

1. **NO usar bypass SSL** en producción
2. **Rotar credenciales** cada 90 días
3. **Auditar accesos** a certificados
4. **Logs no deben contener** tokens o passwords
5. **Cifrar archivos de logs** en reposo

### Cumplimiento

- ✅ GDPR compliant (no almacena datos personales sensibles)
- ✅ Logs auditables en SharePoint
- ✅ Autenticación OAuth 2.0
- ✅ TLS 1.2+ para todas las comunicaciones

---

**FIN DEL DOCUMENTO**

---

*Este documento fue generado automáticamente basándose en el código fuente del proyecto Metadata Manager Process.*
