# Gigas-Panel-Workspace
Panel de control de WorkSpace
# 🛡️ Google Workspace Admin Panel

> **Gestión centralizada y segura de identidades, grupos y procesos críticos de offboarding para Google Workspace mediante Google Apps Script.**

Este proyecto es una aplicación web interna construida sobre **Google Apps Script** que interactúa de forma nativa con las APIs de administración de Google Workspace. Su objetivo es proporcionar a los operadores y mánagers una interfaz gráfica (GUI) simplificada para gestionar usuarios y datos corporativos, sin necesidad de otorgarles privilegios completos de Superadministrador en la consola oficial de Google.

---

## ✨ Características Principales

* **👥 Creación y Gestión de Identidades:** Interfaz unificada para dar de alta nuevos empleados, asignar licencias y configurar datos organizativos de forma estandarizada.
* **📦 Traspaso de Propiedad de Drive (TakeOut/Herencia):** Motor de transferencia directa que utiliza `UrlFetchApp` para comunicarse con la API nativa de transferencias de Google (`datatransfer/v1`). Permite mover todos los archivos de un empleado saliente al Drive de su mánager sin descargar la información al equipo local.
* **🛑 Acciones Críticas de Bajas (Offboarding):** Módulos seguros para suspender cuentas de acceso, restaurar usuarios eliminados accidentalmente (dentro del periodo de retención legal) o purgar identidades definitivamente.
* **📋 Gestión de Listas y Grupos:** Visualización, limpieza y mantenimiento de listas de distribución corporativas.
* **🔐 Sistema de Permisos Dinámico (Whitelist):** El panel cuenta con un sistema de acceso tipo "Toggle" restringido. Solo el Superadministrador maestro puede autorizar o revocar accesos a otros operadores introduciendo su correo en un pop-up, lo que actualiza de forma automática y silenciosa una base de datos segura.

---

## 🛠️ Stack Tecnológico

* **Backend:** Google Apps Script (`.gs`, JavaScript basado en V8).
* **Frontend:** HTML5, CSS3, Vanilla JavaScript.
* **APIs de Google Integradas:** 
  * `Admin SDK (AdminDirectory)`: Para traducciones de correo a ID numérico interno y gestión de estados.
  * `Data Transfer API`: Para el movimiento de archivos en segundo plano.
  * `SpreadsheetApp`: Como base de datos ligera para registros, logs y listas blancas de seguridad.

---

## 🚀 Despliegue e Instalación

Para implementar este panel en un nuevo entorno de Google Workspace:

1. Crea un nuevo proyecto en [Google Apps Script](https://script.google.com/).
2. Copia los archivos del repositorio manteniendo la estructura (archivos `.gs` para el backend y `.html` para la interfaz).
3. En el editor de Apps Script, ve al menú lateral izquierdo **Servicios (+)** y añade la API de **Admin SDK API** (necesaria para las comprobaciones de directorio).
4. Implementa el proyecto como una **Aplicación Web**:
   * **Ejecutar como:** El usuario que accede a la aplicación web (o tu cuenta de Superadministrador, según la arquitectura de permisos deseada).
   * **Quién tiene acceso:** Cualquier usuario dentro de la organización.
5. Autoriza los permisos de oAuth en la primera ejecución.

### 🌐 Entornos de Ejecución
* **Producción (`/exec`):** URL estable para el uso diario por parte de los operadores.
* **Desarrollo (`/dev`):** URL que refleja los cambios de código guardados en tiempo real, ideal para depurar la interfaz (refrescando con `Ctrl + Shift + R`) sin afectar a los usuarios en producción.

---

## 🛡️ Notas de Seguridad

* **Acceso Externo:** El sistema de lista blanca está diseñado para aceptar dominios externos o secundarios previa autorización manual del Superadministrador, permitiendo la operación por parte de proveedores o departamentos aislados sin comprometer el tenant principal.
* **Prevención de Errores:** Todos los procesos críticos solicitan confirmación previa por interfaz y validan la existencia de cuentas activas/suspendidas antes de lanzar llamadas destructivas a los servidores de Google.
