# Reportes colectivos Scoresway

Aplicacion Streamlit para generar dos reportes visuales de partidos a partir de links de Scoresway y 365Scores.

## Archivos

- `scoresway_streamlit_app.py`: app principal standalone.
- `requirements.txt`: dependencias Python para Streamlit Cloud.
- `packages.txt`: dependencias Linux para Chromium/Selenium.
- `.streamlit/config.toml`: configuracion basica de Streamlit.

Los archivos auxiliares originales (`typeId.xlsx`, `qualifiers.csv`, `xT_Grid.csv`) estan embebidos dentro del `.py`, igual que el motor del notebook.

## Uso local

```bash
pip install -r requirements.txt
streamlit run scoresway_streamlit_app.py
```

## Deploy en Streamlit Community Cloud

1. Subir esta carpeta a un repo de GitHub.
2. Crear una app en Streamlit Cloud.
3. Elegir el repo, branch y archivo principal `scoresway_streamlit_app.py`.
4. En `Advanced settings`, elegir Python 3.12.

Nota: la app actualmente conserva el flujo de captura de endpoints del notebook original. En entornos cloud, si Selenium da problemas, el siguiente paso recomendado es reemplazar esa captura por llamadas directas a los endpoints de Scoresway.

Si el build queda mucho tiempo en `Preparing metadata (pyproject.toml)`, suele ser por la instalacion/resolucion de dependencias cientificas pesadas (`pandas`, `pillow`, `scipy`, `scikit-learn`, `matplotlib`) o por dependencias transitivas. En Streamlit Cloud conviene usar Python 3.12 para favorecer wheels precompilados. No uses `runtime.txt`: Streamlit Community Cloud no lo usa para seleccionar la version de Python; se configura desde `Advanced settings`.
