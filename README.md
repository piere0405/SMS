# Generador Inteligente de Campañas de Activación — BBVA / PlusMetas

App Streamlit (`app.py`). Cruza FORMALIZADAS vs ACTIVADAS por DNI, segmenta por tipo de
tarjeta, toma los primeros 999 NO activados, analiza las promociones del Arrancón (Ago-2026),
da alertas + recomendación por promoción y genera el Excel `REC OTROS` (Telefono | Mensaje | DNI).

## Ejecutar
    pip install -r requirements.txt
    streamlit run app.py

## Lógica clave
- FORMALIZADAS: hojas `OUT` + `Hoja2` de BBVA-TLM. ACTIVADAS: hoja `DTA JULIO`.
- Cruce por DNI -> excluye activados -> valida -> dedup (conserva orden) -> primeros 999.
- Nombre: formato BBVA `APELLIDO APELLIDO NOMBRE` (toma el primer nombre real). Sin opción de orden.
- Teléfono de salida: prefijo `51` (51999999999).
- Segmento `Tarjeta_WF`: VISA CERO / Otros / Todos. Las promos de **Pagos Sin Intereses**
  (Educación, Deporte) **no se envían a VISA CERO**; en modo "Todos" se excluyen los cero
  automáticamente si se elige una campaña PSI.
- Mensaje: `NOMBRE, <speech>. activa hoy por la app BBVA o cajero mas cercano` · sin tildes · <=160.
- Paso 4: alertas de vencimiento + recomendación de cuándo usar cada promoción (💡 hipótesis /
  🎯 recomendación; sin cifras inventadas). Ya no hay tarjeta única de "recomendación grande".
- Excel: formato idéntico a la plantilla `sms` (hoja REC OTROS, cabecera plana, Telefono numérico,
  DNI texto, anchos A=12.43 / B=124.29 / C=11.57).
- Promoción "Tarjeta de Crédito Start": ELIMINADA.
