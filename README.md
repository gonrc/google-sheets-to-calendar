# 🗓️ Agregador de exámenes finales a Google Calendar

Este script está pensado para estudiantes que reciben sus fechas de exámenes en un Google Sheets (o similar) y quieren agregarlas rápidamente a Google Calendar.

También puede adaptarse a otros usos donde sea útil cargar automáticamente múltiples eventos a Google Calendar desde una hoja de cálculo (reuniones, turnos, clases, etc).

---

## 🔎 ¿Qué hace este script?

- Lee un rango de filas en una hoja de Google Sheets.
- Interpreta las fechas y horas (en formato como “martes, 16 de julio de 2025” y “14:00”).
- Crea los eventos correspondientes en Google Calendar.
- Evita crear duplicados si el evento ya existe.
- Agrega recordatorios (email y/o popup) según tu configuración.

---

## 📋 Cómo debe estar organizada la hoja

El script lee el rango que definas en la configuración. Cada fila dentro del rango debería seguir esta estructura:

| Código | Materia                      | Fecha 1° llamado             | Hora 1° | Fecha 2° llamado             | Hora 2° |
|--------|------------------------------|------------------------------|---------|------------------------------|---------|
| 11.85  | Toma de Decisiones           | jueves, 17 de julio de 2025 | 13:00   | jueves, 24 de julio de 2025 | 13:00   |
| 71.90  | Certificaciones tecnológicas | viernes, 4 de julio de 2025 | 8:00    | martes, 15 de julio de 2025 | 8:00    |
| ...    | ...                          | ...                          | ...     | ...                          | ...     |

- El código es opcional y no se usa.
- El nombre de la materia se usa como título del evento.
- El segundo llamado también es opcional.
- Si hay filas vacías o incompletas, se omiten automáticamente.

---

## 🛠️ Cómo usarlo

1. Abrí [Google Apps Script](https://script.google.com/) y creá un nuevo proyecto.
2. Pegá el contenido del archivo `.gs` dentro del editor.
3. Configurá los parámetros de la sección `CONFIG` al principio del script
4. Ejecutá la función `addFinalExamsToCalendar`.
5. Revisá el registro (Menú → Ver → Registros) para ver qué se creó o si hubo errores.

---

## ⚙️ Parámetros personalizables

Se pueden modificar desde el bloque `CONFIG` al inicio del script:

| Parámetro                | Qué hace                                                                 |
|--------------------------|---------------------------------------------------------------------------|
| `spreadsheetId`          | ID de la planilla de Google Sheets desde donde se leerán los datos.      |
| `sheetName`              | Nombre de la pestaña (hoja) dentro de la planilla.                       |
| `dataRange`              | Rango de celdas que contiene los datos (ej. `"A2:F9"`).                   |
| `calendar`               | Calendario de Google donde se crearán los eventos.                       |
| `eventDurationHours`     | Duración de cada evento en horas.                                        |
| `rowDelayMs`             | Tiempo de espera entre procesar una fila y otra (en milisegundos).       |
| `apiRateLimitDelayMs`    | Espera (en milisegundos) si Google Calendar lanza un error por exceso.   |
| `avoidDuplicates`        | Si está en `true`, no se crean eventos si ya existen con mismo nombre y hora. |
| `emailRemindersDays`     | Lista de días antes del evento para enviar recordatorios por email.      |
| `popupRemindersDays`     | Lista de días antes del evento para mostrar recordatorios emergentes.    |

---

## 📌 Notas adicionales

- El formato de las fechas debe estar en español, con el mes en minúsculas.
- Si configurás un rango más grande que la cantidad de datos reales, no pasa nada.
- Todos los eventos se crean en el calendario que elijas en la configuración (`getDefaultCalendar()` o cualquier otro).
- Si Google Calendar devuelve errores por uso excesivo, el script espera automáticamente antes de continuar.

---

## 👤 Autor

Gonzalo Ruiz Camauer  
✉️ recipes_ficus_0s@icloud.com
