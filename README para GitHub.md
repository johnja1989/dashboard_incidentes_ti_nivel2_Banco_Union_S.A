# 📊 Dashboard Backlog Incidentes TI

Dashboard estadístico en tiempo real para la gestión y visualización de incidentes de TI. Sistema diseñado para migrar datos operacionales desde archivos Excel hacia visualizaciones interactivas y dinámicas.

## 🎯 Características Principales

- ✅ **Visualización en Tiempo Real**: Gráficos interactivos actualizados automáticamente
- 📈 **Múltiples Métricas**: KPIs, estados, tiempos, responsables, servicios y proveedores
- 💾 **Persistencia Local**: Los datos se mantienen al refrescar la página
- 📱 **100% Responsive**: Funciona en cualquier dispositivo
- 🖨️ **Reportes Profesionales**: PDF ejecutivo y Excel detallado
- 📊 **Múltiples Hojas Excel**: Resumen, datos, rankings y análisis
- 🔄 **Actualización Simple**: Solo requiere cargar un CSV

## 🚀 Demo en Vivo

👉 **[Ver Dashboard](https://tu-usuario.github.io/dashboard-incidentes)**

## 📋 Requisitos

- Navegador moderno (Chrome, Firefox, Safari, Edge)
- JavaScript habilitado
- Archivo CSV con datos de incidentes

## 🛠️ Tecnologías Utilizadas

| Componente | Tecnología | Versión |
|------------|------------|---------|
| Frontend | HTML5 + CSS3 + JavaScript | Nativo |
| Gráficos | Chart.js | 4.4.0 |
| Procesamiento CSV | PapaParse | 5.4.1 |
| Etiquetas en gráficos | ChartDataLabels | 2.2.0 |
| Generación PDF | jsPDF + AutoTable | 2.5.1 / 3.5.31 |
| Exportación Excel | SheetJS (xlsx) | 0.18.5 |
| Hosting | GitHub Pages | - |

## 📂 Estructura del Proyecto

```
dashboard-incidentes/
├── index.html          # Página principal (copia de dashboard.html)
├── dashboard.html      # Dashboard principal
├── app.js             # Lógica de la aplicación
├── styles.css         # Estilos
├── datos.csv          # Datos de ejemplo (opcional)
└── README.md          # Documentación
```

## 🔧 Instalación y Configuración

### Opción 1: Usar directamente en GitHub Pages

1. **Fork o crea un nuevo repositorio**
2. **Sube los archivos** (dashboard.html, app.js, styles.css)
3. **Renombra** `dashboard.html` a `index.html` (o crea una copia)
4. **Activa GitHub Pages**:
   - Ve a `Settings` > `Pages`
   - En `Source` selecciona `main` branch
   - Guarda los cambios
5. **Accede** a `https://tu-usuario.github.io/nombre-repo`

### Opción 2: Clonar y ejecutar localmente

```bash
# Clonar el repositorio
git clone https://github.com/tu-usuario/dashboard-incidentes.git

# Entrar al directorio
cd dashboard-incidentes

# Abrir con un servidor local (ej. con Python)
python -m http.server 8000

# O con Node.js
npx http-server
```

Luego abre `http://localhost:8000` en tu navegador.

## 📥 Formato del CSV

El dashboard espera un archivo CSV con codificación **Windows-1252** (CP1252). Columnas principales esperadas:

```csv
Estado Final Incidente,Ingeniero Asignado,Edad Incidente,Rango_Edad,Servicio,Proveedor a escalar
Abierto,Juan Pérez,15,15-30 días,Base de Datos,Oracle
Cerrado,María López,5,< 7 días,Aplicación,Microsoft
```

### Columnas Reconocidas Automáticamente

El sistema detecta automáticamente estas columnas (no importa el orden):

- **Estado**: `Estado Final Incidente`, `Estado`, `Status`
- **Responsable**: `Ingeniero Asignado`, `Responsable`, `Asignado`
- **Tiempo**: `Edad Incidente`, `Días`, `Tiempo`, `Duración`
- **Rango de Edad**: `Rango_Edad`, `Rango Edad`
- **Servicio**: `Servicio`, `Tipificación`, `Categoría`
- **Proveedor**: `Proveedor a escalar`, `Proveedor`, `Vendor`

## 🎨 Personalización

### Cambiar Colores

Edita las variables CSS en `styles.css`:

```css
:root {
  --bg: #0f172a;
  --card: #111827;
  --text: #e5e7eb;
  --accent: #4cc9f0;
  --accent2: #219ebc;
}
```

### Ajustar Alturas de Gráficos

```css
:root {
  --height-kpi: 360px;
  --height-chart: 320px;
  --height-table: 520px;
}
```

## 📊 Uso Diario

### Actualizar Datos

1. **Exporta tu Excel a CSV**
   - Archivo > Guardar como > CSV (delimitado por comas)
   - Codificación: Windows-1252

2. **Carga el archivo**
   - Click en "📂 Cargar CSV"
   - Selecciona el archivo
   - Los datos se actualizan automáticamente

3. **Verificación**
   - Los datos quedan guardados en el navegador
   - Al refrescar la página, se mantienen
   - Click en "🧹 Limpiar caché" para resetear

### Generar Reportes

#### Reporte PDF Ejecutivo

1. Click en "📊 Reporte PDF Ejecutivo"
2. Se descarga automáticamente un PDF con:
   - Resumen ejecutivo con KPIs principales
   - Top 5 responsables con más casos
   - Top 5 servicios con más incidentes
   - Distribución por rango de edad
   - Diseño profesional con gráficos y tablas

#### Exportar a Excel

1. Click en "📗 Exportar Excel"
2. Se descarga un archivo Excel con 6 hojas:
   - **Resumen Ejecutivo**: Métricas principales
   - **Datos Completos**: Toda la información cargada
   - **Top Responsables**: Ranking de casos por ingeniero
   - **Top Servicios**: Ranking de incidentes por servicio
   - **Análisis por Rango**: Tiempos promedio por edad
   - **Proveedores**: Casos escalados por proveedor

> 💡 **Tip**: Usa el PDF para presentaciones rápidas a gerencia y el Excel para análisis detallado.

## 🔐 Seguridad

- ✅ Todo el procesamiento es local (en el navegador)
- ✅ No se envían datos a servidores externos
- ✅ Cifrado HTTPS en GitHub Pages
- ✅ Control de acceso mediante configuración del repositorio

### Hacer el Repositorio Privado

1. Ve a `Settings` del repositorio
2. Scroll hasta "Danger Zone"
3. Click en "Change visibility" > "Make private"

> **Nota**: GitHub Pages en repos privados requiere GitHub Pro

## 📈 Métricas Mostradas

### KPIs Principales
- **Total de Incidentes Reportados**
- **Abiertos vs Cerrados**
- **Tiempo Promedio de Resolución**

### Gráficos Disponibles
1. **Estado de Incidentes** (Torta)
2. **Tiempo por Rango de Edad** (Barras)
3. **Casos por Responsable** (Barras)
4. **Tipificación por Servicio** (Barras)
5. **Casos por Proveedor** (Barras)
6. **Tabla de Datos Detallada**

## 🐛 Solución de Problemas

### El dashboard no carga

- Verifica que GitHub Pages esté activado
- Asegúrate de que el archivo se llame `index.html`
- Revisa la consola del navegador (F12)

### Los acentos se ven mal

- El archivo CSV debe estar en codificación **Windows-1252**
- Si exportas desde Excel, usa "CSV (delimitado por comas)"

### Los gráficos no se actualizan

- Limpia el caché del navegador
- Click en "🧹 Limpiar caché" en el dashboard
- Recarga la página (Ctrl+F5)

### "Sin datos" en los gráficos

- Verifica que el CSV tenga las columnas correctas
- Revisa la consola (F12) para ver errores
- Asegúrate de que los nombres de columnas coincidan

## 👨‍💻 Autor

**John Jairo Vargas González**  
Ingeniero de Soluciones TI  
📧 john.vargas@bancounion.com

---

## 📄 Licencia

Este proyecto es de uso interno corporativo.

## 🤝 Contribuciones

Para sugerencias o mejoras, contacta al autor.

---

**"Transformando ideas en soluciones tecnológicas"** ✨