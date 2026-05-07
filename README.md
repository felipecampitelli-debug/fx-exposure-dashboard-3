# FX Exposure Dashboard - Streamlit

## 🚀 Instalación

1. Instala las dependencias:
```bash
pip install -r requirements.txt
```

## ▶️ Ejecutar la App

```bash
streamlit run fx_dashboard_app.py
```

Se abrirá automáticamente en tu navegador: http://localhost:8501

## 📋 Uso

1. **Carga los 7 archivos** requeridos en el panel lateral
2. **Presiona "GENERAR ANÁLISIS"**
3. **Visualiza** las métricas interactivas
4. **Descarga** el Excel generado

## 📁 Archivos Requeridos

- `gl_balance_de_comprobacion_DDMMAA.txt` - Mayor contable
- `ccy_spot_latam.xlsx` - Tipos de cambio
- `ndf_position*.xlsx` - Posiciones NDF
- `ops_no_capturadas*.xlsx` - Ops no capturadas
- `flex_taxes.xlsx` - Flex Taxes
- `vtos_prestamos_chile.xlsx` - Préstamos Chile
- `fx_exposure_YYYYMMDD.xlsx` - Histórico

## 🎯 Características

✅ Procesamiento 100% local
✅ Dashboard interactivo en navegador
✅ Gráficos dinámicos con Plotly
✅ Descarga de Excel completo
✅ KPIs en tiempo real
✅ Cronograma de NDFs
