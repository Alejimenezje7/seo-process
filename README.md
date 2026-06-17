# Mapa de Procesos SEO & ASO

Aplicacion web interna para documentar y consultar los procesos de SEO tecnico y ASO del equipo LAM.

## Ejecutar localmente

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Deploy en Streamlit Community Cloud

1. Sube este repo a GitHub
2. Ve a [share.streamlit.io](https://share.streamlit.io)
3. Conecta tu cuenta de GitHub
4. Selecciona este repo → `app.py`
5. Click en **Deploy**

La app estara disponible en una URL tipo `https://tu-app.streamlit.app`

## Acceso

La app esta protegida con contrasena. Contacta al administrador del equipo para obtener las credenciales.

## Estructura

```
app.py                  # Aplicacion principal
requirements.txt        # Dependencias Python
.streamlit/config.toml  # Tema oscuro personalizado
```

## Contenido

### SEO
- Redirecciones 404
- Estructura de URLs
- Mobile Optimization
- Rendimiento
- Metadata & Schema
- Indexacion

### ASO
- ASO Optimizations
- In App Events
- AB Test
