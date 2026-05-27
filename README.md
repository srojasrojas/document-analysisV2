# DOCX rule engine

Motor simple para editar `.docx` mediante reglas configurables de reemplazo o
expansion. Las reglas iniciales amplian menciones de `operador` y `supervisor`,
pero el codigo no depende de esos cargos.

El proyecto evita la fase pesada de OCR/chunking del repositorio anterior. La
unidad de trabajo es el parrafo editable de Word, incluyendo parrafos dentro de
celdas de tablas.

## Configuracion

La configuracion principal vive en `config.yaml`:

- `rules`: detectores, frase objetivo, formato de reemplazo, guardas y prompts.
- `pipeline.max_passes`: numero maximo de pasadas.
- `llm_refine.model`: modelo usado por el refining LLM.
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

Crear el mock `.docx` editable multi-regla:

```powershell
venv\Scripts\python.exe scripts\create_mock_docx.py
```

Correr el flujo sobre el mock:

```powershell
venv\Scripts\python.exe -m rule_engine.pipeline --config config.yaml --input data\input\mock_reglas.docx --output data\output --passes 3 --force
```

En maquinas sin el modelo local `deepseek-r1-distill-14b-local`, se puede validar
solo la capa deterministica sin invocar LLM:

```powershell
venv\Scripts\python.exe -m rule_engine.pipeline --config config.yaml --input data\input\mock_reglas.docx --output data\output --passes 3 --force --simple-only
```

Verificar que no haya duplicados de ninguna frase objetivo configurada:

```powershell
venv\Scripts\python.exe scripts\verify_output.py data\output\mock_reglas_modificado.docx --config config.yaml
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

El refining LLM queda habilitado por defecto con:

```yaml
llm_refine:
	model: deepseek-r1-distill-14b-local
```

Ese perfil usa `provider: ollama`, `deployment: deepseek-r1:14b` y
`base_url: http://192.168.0.24:11434/v1`, siguiendo la configuracion actual del
repo vecino. Para usar OpenAI directo, cambia `llm_refine.model` a `openai_api`.

## Salidas

- `data/input/mock_reglas.docx`: documento mock editable.
- `data/output/mock_reglas_modificado.docx`: documento procesado.
- `reports/changes.jsonl`: cambios aplicados, con ubicacion body/tabla.
- `reports/registro_cambios.xlsx`: registro legible con cambios y omitidos.
- `reports/referencia_resumen.md`: resumen del Excel manual.
