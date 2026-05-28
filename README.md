# DOCX rule engine

Motor simple para editar `.docx` mediante reglas configurables de reemplazo o
expansion. La rama `no-llm` deja el flujo en modo deterministico: no invoca LLM
para detectar, seleccionar ni redactar cambios.

El proyecto evita la fase pesada de OCR/chunking del repositorio anterior. La
unidad de trabajo es el parrafo editable de Word, incluyendo parrafos dentro de
celdas de tablas.

## Configuracion

La configuracion principal vive en `config.yaml`:

- `rules`: detectores, frase objetivo, formato de reemplazo, guardas y prompts.
- `pipeline.max_passes`: numero maximo de pasadas.
- `pipeline.use_llm_refine`: en esta rama queda en `false`.
- `pipeline.post_audit`: auditoria final y reparacion segura de cambios obvios.
- `llm_refine.model`: perfil disponible, pero no usado mientras `use_llm_refine`
	sea `false`.
- `models.deepseek-r1-distill-14b-local`: perfil local Ollama/OpenAI-compatible,
  configurado igual que en el repo vecino `document-analysis`.
- `models.openai_api`: perfil alternativo para API OpenAI con `OPENAI_API_KEY`.

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

Extraer el corpus de casos de operador desde la corrida `tmp/run_b25`:

```powershell
venv\Scripts\python.exe scripts\extract_operator_corpus.py --run-dir tmp\run_b25 --output reports\operator_corpus.xlsx
```

Crear el mock `.docx` editable multi-regla:

```powershell
venv\Scripts\python.exe scripts\create_mock_docx.py
```

Correr el flujo deterministico sobre el mock:

```powershell
venv\Scripts\python.exe -m rule_engine.pipeline --config config.yaml --input data\input\mock_reglas.docx --output data\output --passes 3 --force --simple-only
```

Correr sobre una carpeta de documentos:

```powershell
venv\Scripts\python.exe -m rule_engine.pipeline --config config.yaml --input tmp\run_b25 --output tmp\run_b25_no_llm_reviewed --passes 3 --force --simple-only
```

Verificar que no haya duplicados de ninguna frase objetivo configurada:

```powershell
venv\Scripts\python.exe scripts\verify_output.py data\output\mock_reglas_modificado.docx --config config.yaml
```

Verificar todos los documentos de una corrida:

```powershell
$docs = @(Get-ChildItem tmp\run_b25_no_llm_reviewed -Filter *.docx); $failures = @(); foreach ($doc in $docs) { & venv\Scripts\python.exe scripts\verify_output.py $doc.FullName --config config.yaml | Out-Null; if ($LASTEXITCODE -ne 0) { $failures += $doc.Name } }; $failures
```

Validar idempotencia sobre la salida revisada:

```powershell
venv\Scripts\python.exe -m rule_engine.pipeline --config config.yaml --input tmp\run_b25_no_llm_reviewed --output tmp\dry_run_unused --passes 3 --simple-only --dry-run
```

La misma corrida puede repetirse. Si `data\output\mock_reglas_modificado.docx`
ya existe, el pipeline continua desde ese archivo y no vuelve a copiar el input,
lo que permite validar idempotencia.

## Agregar reglas

Cada entrada de `rules` contiene:

- `detection.regex`: patron para encontrar candidatos en parrafos y tablas.
- `replacement.target_phrase`: texto a insertar o reemplazar.
- `replacement.format`: template, por ejemplo `{matched} {connector} {target}`.
- `guards`: excepciones, alternativas ya existentes y textos solo para revision.
- `llm_refining`: prompts especificos y validaciones opcionales por regla.

La regla de operador usa dos destinos:

- Por defecto agrega `o personal designado por minera Spence`.
- Usa `o personal calificado` cuando el contexto menciona personal autorizado,
  capacitado, certificado, habilitado, bloqueo/LOTO, energizacion, HMI, panel de
  control, reset, instrumentacion, mantencion, calibracion o diagnostico tecnico.

Las menciones bajo secciones `REGISTRO` o `REGISTROS` se omiten y quedan en el
registro como `skip_section`. Tambien se omiten contextos CAS/CIO/Sala de
Control, alternativas existentes y celdas/titulos sin verbo de accion.

El perfil LLM queda documentado, pero deshabilitado en esta rama mientras
`pipeline.use_llm_refine` y `rules[*].llm_refining.enabled` sigan en `false`:

```yaml
llm_refine:
	model: deepseek-r1-distill-14b-local
```

Ese perfil usa `provider: ollama`, `deployment: deepseek-r1:14b` y
`base_url: http://192.168.0.24:11434/v1`, siguiendo la configuracion actual del
repo vecino. Para usar OpenAI directo en otra rama, cambia `llm_refine.model` a
`openai_api` y habilita explicitamente el refining.

## Salidas

- `data/input/mock_reglas.docx`: documento mock editable.
- `data/output/mock_reglas_modificado.docx`: documento procesado.
- `tmp/run_b25_no_llm_reviewed/*.docx`: documentos reales procesados en modo no-LLM.
- `reports/operator_corpus.xlsx`: coleccion de candidatos extraida de `run_b25`.
- `reports/changes.jsonl`: cambios aplicados, con ubicacion body/tabla.
- `reports/registro_cambios.xlsx`: registro legible con cambios y omitidos.
- `reports/auditoria_post_run.xlsx`: auditoria de cambios, QA flags y reparaciones.
- `reports/auditoria_post_run.json`: misma auditoria en formato estructurado.
- `reports/referencia_resumen.md`: resumen del Excel manual.

La corrida puede repetirse sobre el mismo output. Si el documento ya esta
expandido, debe terminar con `0 change(s)` y conservar los conteos sin duplicar
targets.
