# Creador IG · MÖÖnna

App para crear Reels de Instagram con voz en off, pensada para usarse desde el iPhone.

```
docs/            App web (GitHub Pages). Diseño MÖÖnna. Graba con el micrófono y habla con el backend.
apps-script/     Backend en Google Apps Script: ideas y transcripción con Gemini, datos en Sheets + Drive.
```

La PC (proyecto `Creador_IG`, carpeta vecina) queda como **editor de video**: cuando está encendida,
toma las grabaciones de Drive y arma los Reels con FFmpeg. (Paso 3, pendiente.)

## Puesta en marcha

### 1. Backend (Apps Script)
1. `npm install -g @google/clasp` y `clasp login` (abre el navegador; inicia sesión con tu Google).
2. En `apps-script/`: `clasp create --type webapp --title "Creador IG · MÖÖnna"` y luego `clasp push`.
3. `clasp open-script` para abrir el editor:
   - **Configuración del proyecto → Propiedades del script**: agrega `GEMINI_API_KEY`.
   - Ejecuta la función **`instalar`** y acepta los permisos. En el registro aparece el **CÓDIGO**.
   - **Implementar → Nueva implementación → Aplicación web**, ejecutar como *yo*, acceso *cualquier usuario*.
     Copia la **URL /exec**.

### 2. App web (GitHub Pages)
1. Sube este repositorio a GitHub y activa Pages desde la carpeta `/docs` de la rama `main`.
2. Abre la URL de Pages en el iPhone, pega la URL /exec y el CÓDIGO, y dale a **Conectar**.
3. Compartir → **Agregar a pantalla de inicio**.

## Notas
- Los datos viven en la hoja "Creador IG · MÖÖnna" de tu Drive. El perfil de marca que lee la IA está
  en la pestaña `marca`: edítalo cuando quieras.
- La URL /exec no cambia aunque subas código nuevo con `clasp push`; solo hay que crear una
  **nueva versión** de la implementación (Implementar → Administrar implementaciones → editar → versión nueva).
