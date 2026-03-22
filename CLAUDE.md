# EXCEL_ANALYZER — Claude Configuration

## Proyecto
Librería Python para parsear archivos Excel no estructurados usando openpyxl.
Arquitectura basada en pipeline: scan → detect → parse → normalize → build.

## Skills globales — leer en TODA tarea

Leer estos tres archivos al inicio de cada conversación, sin excepción:

- `C:\Users\AleJi\OneDrive\Documentos\Claude\Skills\Claude_optimization\memory.md`
- `C:\Users\AleJi\OneDrive\Documentos\Claude\Skills\Claude_optimization\pattern-learning.md`
- `C:\Users\AleJi\OneDrive\Documentos\Claude\Skills\Claude_optimization\execution-planning.md`

## Skills específicas — leer cuando aplica

### Cualquier cambio de código
`C:\Users\AleJi\OneDrive\Documentos\Claude\Skills\Improve code\develope_code.md`

### Refactoring o mejora de arquitectura
`C:\Users\AleJi\OneDrive\Documentos\Claude\Skills\Improve code\SKILL.md`
`C:\Users\AleJi\OneDrive\Documentos\Claude\Skills\Improve code\REFERENCE.md`

## Reglas del proyecto

- No romper interfaces existentes sin avisar
- Todo fix debe ser heurística general, no parche específico
- Renombrar columnas siempre POST-scrape, nunca antes
