# 🍧 Sistema de Gestión de Cepillados

![Versión](https://img.shields.io/badge/versión-1.1-blue)
![Python](https://img.shields.io/badge/Python-3.8+-green)
![Flask](https://img.shields.io/badge/Flask-3.0-red)
![Bootstrap](https://img.shields.io/badge/Bootstrap-5.3-purple)
![Licencia](https://img.shields.io/badge/Licencia-MIT-yellow)

Sistema administrativo completo para control de inventario, pedidos y ventas de cepillados. Diseñado para pequeñas y medianas empresas que necesitan gestionar sus ventas diarias de forma eficiente, con facturación profesional y panel de configuraciones integrado.

---

## 📸 Capturas de Pantalla

### Dashboard Principal
*Dashboard con estadísticas en tiempo real y despacho rápido de pedidos*

### Panel de Reportes
*Gráficos interactivos de ventas, métodos de pago y tendencias*

### Gestión de Inventario
*Control de stock con imágenes de sabores para publicidad*

### Factura Profesional
*Facturas personalizables con número secuencial y datos de la empresa*

---

## 🚀 Características Principales

### 📊 Dashboard
- Estadísticas en tiempo real (pedidos del día, recaudado, pendientes)
- Despacho rápido de pedidos con número de factura automático
- Últimas transacciones con acciones rápidas (editar, factura, eliminar)
- Indicadores de stock bajo con alertas visuales

### 🧾 Facturación Profesional
- **Número de factura secuencial** (FACT-2026-0001, FACT-2026-0002...)
- Encabezado con datos de la empresa personalizables
- Tabla de productos con precios y totales
- Equivalente en Bolívares según tasa BCV
- Pie de página con mensaje personalizable
- Colores de factura configurables desde el panel
- **Sin IVA** - Total directo y claro

### ⚙️ Panel de Configuraciones
- **Datos de la Empresa**: Nombre, RIF, teléfono, dirección, email
- **Mensaje de factura**: Texto personalizable al final de cada factura
- **Tasa BCV**: 
  - Fuente automática (BCV) o manual
  - Tiempo de caché configurable
  - Opción de mostrar/ocultar en facturas
- **Facturas**: Colores principal y secundario personalizables
- **App**: Nombre del sistema, registros por página

### 📝 Gestión de Pedidos
- Registro de pedidos con nombre, área, sabor y cantidad
- Múltiples métodos de pago configurables:
  - 📱 Pago Móvil (con comprobante)
  - 💵 Efectivo Bolívares
  - 💲 Efectivo Dólares
  - 🏦 Transferencia Bancaria
  - 💸 Zelle
  - Y más... (personalizables)
- Captura de comprobantes de pago
- Control de entregas (pendiente/entregado)
- Edición y eliminación de pedidos con restauración de stock

### 📦 Inventario
- Gestión completa de sabores (CRUD)
- Control de stock con alertas visuales:
  - 🟢 Stock alto (>20)
  - 🟡 Stock medio (11-20)
  - 🔴 Stock bajo (1-10)
  - ⚫ Agotado (0)
- Precios configurables por sabor
- **📸 Imágenes de sabores** para publicidad y referencia visual
- Visualización de imágenes en grande con click
- Gestión de áreas organizacionales
- Gestión de métodos de pago (activar/desactivar/crear/eliminar)

### 📈 Reportes y Estadísticas
- Gráficos de ventas diarias (últimos 30 días)
- Distribución de ventas por método de pago (gráfico de dona)
- Top 5 sabores más vendidos con ranking visual
- Top 5 áreas de mayor consumo
- Tasa de pago con barra de progreso
- Métricas generales (total pedidos, facturado, pendientes)
- Exportación a **Excel** con formato profesional
- Exportación a **PDF** con resumen ejecutivo

### 🎨 Interfaz
- **🌙 Modo claro/oscuro** con persistencia en localStorage
- **📱 Diseño totalmente responsive**:
  - ✅ Escritorio (1200px+)
  - ✅ Tablet (768px-1199px)
  - ✅ Móvil grande (576px-767px)
  - ✅ Móvil pequeño (<576px)
- Sidebar colapsable en dispositivos móviles
- Interfaz moderna con Bootstrap 5.3
- Iconografía con Bootstrap Icons
- Animaciones y transiciones suaves
- Tarjetas estadísticas con hover effects

### 💱 Tasa BCV Inteligente
- Consulta automática al Banco Central de Venezuela
- Sistema de caché configurable (por defecto 60 minutos)
- Modo manual para cuando el BCV no responde
- Forzar actualización inmediata
- Última tasa guardada como respaldo
- Indicador de tiempo desde última actualización

---

## 📋 Requisitos

- **Python** 3.8 o superior
- **pip** (gestor de paquetes de Python)
- Conexión a internet (para tasa BCV automática)
- Navegador web moderno (Chrome, Firefox, Edge, Safari)
