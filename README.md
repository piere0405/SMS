# Generador Inteligente de Campañas de Activación — BBVA / PlusMetas

Streamlit app (carga manual). Cruza FORMALIZADAS vs ACTIVADAS por DNI, segmenta por
tarjeta y por antigüedad de formalización, genera el Excel `REC OTROS` (formato tabla).

## Ejecutar
    pip install -r requirements.txt
    streamlit run app.py

## Flujo
1. Subir FORMALIZADAS (BBVA-TLM: hojas OUT+Hoja2) y ACTIVADAS (hoja DTA JULIO).
2. Cruce por DNI -> excluye activados -> valida -> dedup (conserva orden).
3. **Segmento de tarjeta** (Tarjeta_WF): Todos / VISA CERO / Otros.
4. **Segmento por antigüedad** (Fecha_Formalizacion vs hoy): 0-3 / 4-15 / 16-30 / 31+ días.
   Se elige el grupo a enviar y un máximo de registros (default 999).
5. Análisis de promociones + alertas + recomendación por promoción.
6. Selección de campaña -> speech <=160 sin tildes -> vista previa -> Excel.

## Reglas
- Teléfono de salida con prefijo 51 (numérico). DNI como texto. Hoja REC OTROS con formato tabla.
- Nombre: formato BBVA APELLIDO APELLIDO NOMBRE (toma el primer nombre real).
- VISA CERO NO recibe: Puntos y Grandes Premios, Bono Semanal, Educación, x5 Exterior, Deporte.
  Sí recibe: Día del Niño, Exoneración de Membresía, Conciertos.
