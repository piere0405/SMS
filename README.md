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
4. **Tipo de envío** (Tipo_Envio): Todos / Oficina / Courier.
5. **Segmento por antigüedad** (Fecha_Formalizacion vs hoy): 0-3 / 4-15 / 16-30 / 31+ días.
   Se elige el grupo a enviar y un máximo de registros (default 999).
6. Análisis de promociones + alertas + recomendación por promoción.
7. Selección de campaña -> speech <=160 sin tildes -> vista previa -> Excel.

## OTRO (mensaje libre)
En la opción OTRO se escribe TODO el mensaje libremente; solo se antepone el nombre automáticamente (sin forzar el cierre). Se quitan tildes y se valida <=160.

## Reglas
- Teléfono de salida con prefijo 51 (numérico). DNI como texto. Hoja REC OTROS con formato tabla.
- Nombre: formato BBVA APELLIDO APELLIDO NOMBRE (toma el primer nombre real).
- VISA CERO NO recibe: Puntos y Grandes Premios, Bono Semanal, Educación, x5 Exterior, Deporte.
  Sí recibe: Día del Niño, Exoneración de Membresía, Conciertos.

## Actualizaciones recientes
- DNI siempre a 8 dígitos (texto, con ceros a la izquierda).
- La palabra "activa" siempre se muestra como "Activa".
- Módulo 1 Formalizadas: hasta 3 archivos combinados en una sola base.
- Módulo 3 Histórico de Activadas (opcional): hoja ACTIVADAS, DOI, FLAG_ACTIVA=1.
- Cruce excluye DNI activados en Activadas del mes Y/O Histórico (unión, dedup).
- Módulos independientes: funciona con Formalizadas sola, +mes, +histórico o los tres.
- Sección "Recomendaciones" ahora es un desplegable cerrado por defecto.
