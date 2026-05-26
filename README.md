# Operator replacement workflow

Flujo simple para editar `.docx` y ampliar menciones de `operador` hacia
`operador o personal designado por minera Spence` mediante reglas configurables
y pasadas idempotentes.

El proyecto evita la fase pesada de OCR/chunking del repositorio anterior. La
unidad de trabajo es el parrafo editable de Word, incluyendo parrafos dentro de
celdas de tablas.

## Configuracion

La configuracion principal vive en `config.yaml`:

- `rules`: regex, frase objetivo, guardas de idempotencia y casos de revision.
- `pipeline.max_passes`: numero maximo de pasadas.
- `models.active`: perfil activo para refining LLM opcional.
- `models.profiles.openai_api`: usa `OPENAI_API_KEY` y `OPENAI_MODEL`.
- `models.profiles.deepseek-r1-distill-14b-local`: perfil local OpenAI-compatible.

Las credenciales se leen desde variables de entorno. Usa `.env.example` como
plantilla y no guardes claves reales en el repositorio.

## Comandos

Instalar dependencias con el venv del repo:

```powershell
venv\Scripts\python.exe -m pip install -r requirements.txt
```

Inspeccionar el Excel manual de referencia:

```powershell
venv\Scripts\python.exe scripts\inspect_reference.py
```

Crear el mock `.docx` editable:

```powershell
venv\Scripts\python.exe scripts\create_mock_docx.py
```

Correr el flujo sobre el mock:

```powershell
venv\Scripts\python.exe -m operator_replacement.pipeline --config config.yaml --input data\input\mock_operador.docx --output data\output --passes 3
```

Verificar que no haya duplicados de la frase objetivo:

```powershell
venv\Scripts\python.exe scripts\verify_output.py data\output\mock_operador_modificado.docx
```

La misma corrida puede repetirse. Si `data\output\mock_operador_modificado.docx`
ya existe, el pipeline continua desde ese archivo y no vuelve a copiar el input,
lo que permite validar idempotencia.

## Salidas

- `data/input/mock_operador.docx`: documento mock editable.
- `data/output/mock_operador_modificado.docx`: documento procesado.
- `reports/changes.jsonl`: cambios aplicados, con ubicacion body/tabla.
- `reports/registro_cambios.xlsx`: registro legible con cambios y omitidos.
- `reports/referencia_resumen.md`: resumen del Excel manual.
