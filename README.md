# Consulta Masiva de Comprobantes Electrónicos – API SUNAT

Este proyecto permite validar **comprobantes electrónicos de manera masiva** utilizando la **API oficial de la SUNAT**.  
El script lee una plantilla Excel, obtiene el token OAuth2, envía consultas individuales, maneja reintentos automáticos y escribe los resultados en el mismo archivo.

📅 **Versión funcional:** 11/12/2025  

---

## 🚀 Características

- Obtención automática de **token OAuth2 (Client Credentials)**.
- Lectura de parámetros desde una plantilla Excel.
- Validación masiva de:
  - RUC  
  - Tipo de comprobante  
  - Serie  
  - Número  
  - Fecha de emisión  
  - Monto total en moneda original
- Manejo de errores y **reintentos automáticos** cuando la API no responde.
- Omitir filas vacías automáticamente.
- Escritura de resultados en las columnas:
  - **H:** Estado del comprobante  
  - **I:** Estado del RUC  
  - **J:** Condición de domicilio  
  - **K:** Observaciones  
- Apertura automática del archivo Excel al finalizar.

---

## 📂 Estructura esperada de la plantilla Excel

| Celda | Contenido |
|-------|-----------|
| **C3** | RUC del consultante |
| **E3** | Client ID |
| **I3** | Client Secret |
| **B–G (desde fila 7)** | RUC, Tipo, Serie, Número, Fecha, Monto en moneda original|

El script genera la respuesta en las columnas **H–K**.

---

## 🔧 Requisitos

Instalar dependencias:

pip install pandas requests openpyxl

---

## ▶️ Ejecución

Configurar la ruta del archivo Excel:

EXCEL_PATH = r"Ruta de Plantilla"

Ejecutar el script:

python "Validador CPE Masivo - API SUNAT.py"

El archivo Excel se actualizará automáticamente y se abrirá al concluir el proceso.

---

## 🔐 Credenciales SUNAT

El Client ID y Client Secret se obtienen desde el portal de SUNAT.

📘 **Manual oficial (hojas 3–5):**  
https://cpe.sunat.gob.pe/sites/default/files/inline-files/Manual-de-Consulta-Integrada-de-Comprobante-de-Pago-por-ServicioWEB_v2_0.pdf

---

## ⚠️ Notas importantes

- El script solo reintenta las filas con error, sin repetir filas ya procesadas.  
- Se limpia el contenido de **H–K** cuando la fila está vacía.  
- La ejecución finaliza únicamente cuando todas las filas han sido procesadas exitosamente.

---

## 📜 Licencia

Proyecto distribuido bajo la licencia **MIT**.
