# DOCX rule engine

Motor simple para editar `.docx` mediante reglas configurables de reemplazo o
expansion. La rama `no-llm` deja el flujo en modo deterministico: no invoca LLM
para detectar, seleccionar ni redactar cambios.

El proyecto evita la fase pesada de OCR/chunking del repositorio anterior. La
unidad de trabajo es el parrafo editable de Word, incluyendo parrafos dentro de
celdas de tablas.

## Configuracion

La configuracion principal vive en `config.yaml` y puede sobreescribirse de
forma local con `config.local.yaml` sin tocar el repo:

- `rules`: detectores, frase objetivo, formato de reemplazo, guardas y prompts.
- `pipeline.max_passes`: numero maximo de pasadas.
- `pipeline.use_llm_refine`: en esta rama queda en `false`.
- `pipeline.embedded_header_footer_cleanup`: pre-limpieza de encabezados y pies
    falsos incrustados en el cuerpo por conversiones PDF a DOCX.
- `paths.input_dir`: por defecto apunta a `inputs/` para documentos editables
    sueltos.
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

Ese mock se escribe en `inputs/mock_reglas.docx`, que coincide con el input por
defecto del pipeline.

Correr el flujo deterministico sobre el mock:

```powershell
venv\Scripts\python.exe -m rule_engine.pipeline --config config.yaml --input inputs\mock_reglas.docx --output data\output --passes 3 --force --simple-only
```

Si no pasas `--input`, el pipeline toma por defecto la carpeta `inputs/`.

Correr sobre una carpeta de documentos:

```powershell
venv\Scripts\python.exe -m rule_engine.pipeline --config config.yaml --input tmp\run_b25 --output tmp\run_b25_no_llm_reviewed --passes 3 --force --simple-only
```

El pipeline ejecuta antes de las reglas una limpieza de pies/encabezados falsos
incrustados en el cuerpo. La accion por defecto elimina o limpia texto de
parrafos de alta confianza, conserva saltos de pagina/seccion cuando los hay y
deja las tablas repetidas de metadata en revision. Para comparar sin esta capa:

```powershell
venv\Scripts\python.exe -m rule_engine.pipeline --config config.yaml --input tmp\run_b25 --output tmp\run_b25_no_cleanup --passes 3 --force --simple-only --skip-embedded-cleanup
```

Inspeccionar candidatos sin modificar documentos:

```powershell
venv\Scripts\python.exe scripts\inspect_embedded_headers_footers.py --input tmp\run_b25\P-PRPL-OC-506_modificado.docx --output reports\embedded_header_footer_preview.xlsx
```

Aplicar la limpieza sobre copias, manteniendo intacto el input original:

```powershell
venv\Scripts\python.exe scripts\inspect_embedded_headers_footers.py --input tmp\run_b25\P-PRPL-OC-506_modificado.docx --apply --output-dir tmp\oc506_footer_cleanup --output reports\embedded_header_footer_apply.xlsx
```

Si una revision manual confirma que las tablas repetidas tambien son encabezados
convertidos, la utility permite incluirlas con `--remove-tables`; por defecto se
reportan pero no se borran.

Si quieres dejar `tmp/run_b25` como input por defecto solo en tu maquina, crea
un `config.local.yaml` ignorado por git:

```yaml
paths:
    input_dir: tmp/run_b25
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

La regla de operador usa tres destinos, en este orden de prioridad:

- Usa `o personal calificado` cuando el contexto menciona personal autorizado,
    capacitado, habilitado, bloqueo/LOTO, energizacion, HMI, panel de control,
    reset, instrumentacion, mantencion, calibracion o diagnostico tecnico.
- Usa `o personal certificado designado por minera Spence` cuando el cargo de
    operador incluye una especificacion de equipo certificado o pesado, como
    retroexcavadora, excavadora, minicargador, cargador frontal, camion tolva,
    camion pluma, rotopala, apilador, esparcidor, picaroca, puente grua, grua
    horquilla u operador de equipos.
- Por defecto agrega `o personal designado por minera Spence` para operador de
    proceso, planta, area o responsabilidades operacionales generales.

La deteccion de operador tambien cubre descriptores compuestos frecuentes en
tablas, como `Operador Spence`, `Operador EW`, `Operador MLDC`, `Operador de
Patio embarque`, `operador de patio de catodos`, `operadores de otras areas` y
`Operador de la maquina despegadora`, para evitar residuos despues del target.

La regla de supervisor repara el legado `o experto tecnico` y lo homologa a
`o Ejecutivos del Área`.

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

- `inputs/mock_reglas.docx`: documento mock editable.
- `data/output/mock_reglas_modificado.docx`: documento procesado.
- `tmp/run_b25_no_llm_certified_supervisor_fixed/*.docx`: documentos reales
    procesados en modo no-LLM con operadores certificados y supervisor legado
    corregido.
- `tmp/run_b25_no_llm_table_operator_fixed/*.docx`: salida no-LLM con
    descriptores compuestos de operador en tablas corregidos.
- `tmp/run_b25_no_llm_order_fixed/*.docx`: salida previa con correccion de orden
    de descriptores.
- `reports/operator_corpus.xlsx`: coleccion de candidatos extraida de `run_b25`.
- `reports/changes.jsonl`: cambios aplicados, con ubicacion body/tabla.
- `reports/registro_cambios.xlsx`: registro legible con cambios y omitidos.
- `reports/auditoria_post_run.xlsx`: auditoria de cambios, QA flags y reparaciones.
- `reports/auditoria_post_run.json`: misma auditoria en formato estructurado.
- `reports/embedded_header_footer_cleanup.xlsx`: auditoria de pies/encabezados
    falsos detectados, acciones propuestas/aplicadas y candidatos en revision.
- `reports/embedded_header_footer_cleanup.jsonl`: misma auditoria en formato
    estructurado.
- `reports/referencia_resumen.md`: resumen del Excel manual.

La corrida puede repetirse sobre el mismo output. Si el documento ya esta
expandido, debe terminar con `0 change(s)` y conservar los conteos sin duplicar
targets.
