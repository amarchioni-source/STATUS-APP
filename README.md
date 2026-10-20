# Status de producción por venta (SO)

App web: se suben el Excel de **stock**, el Excel de **ventas** y el **nro. de SO**, y devuelve el STATUS con el formato de siempre.

- 🚨 **¡Venta con producto retenido!** si el stock del SO tiene kg con estado distinto de APTO.
- ✅ **Todo OK** si no hay retenido, ni pendientes (🔴), ni excesos sobre el 10% (🟠).
- 🟢 dentro del ±10% · 🔴🚨 en producción (falta más del 10%) · 🟠🚨 excede el about (+10%).
- Difference = Produced − SC (negativo = falta producir, positivo = excede).
- El resultado se ve en pantalla y se imprime desde el botón Imprimir (no genera Excel).
- Los archivos se procesan en memoria y no se guardan.

## Configuración (se edita desde la web de GitHub, sin tocar código)

| Archivo | Para qué |
|---|---|
| `config/columnas.json` | Nombre de las columnas que se toman de cada Excel (ventas y stock). |
| `config/reglas.json` | Tolerancia (0.10), estados que se consideran aptos, prefijos que se quitan de la descripción. |
| `config/equivalencias.json` | Códigos de venta que en producción abren en otros códigos. `"*"` aplica a todos los SO; una clave con el nro. de SO aplica solo a ese. |

Ejemplo de equivalencia: `"10121": {"CD21500A": ["CD216150", "CD216151", "CD216152"]}`

Para editar: abrir el archivo en GitHub → lápiz → Commit changes. Render redespliega solo.

## Publicar

1. Crear un repositorio **privado** en GitHub y subir esta carpeta (los `.xlsx` están excluidos por `.gitignore`: no subir datos comerciales).
2. En Render: **New → Blueprint** y elegir el repositorio (usa `render.yaml`).
3. En **Environment** cargar `APP_USER` y `APP_PASSWORD` (usuario y contraseña de acceso a la app).
4. Al terminar el deploy, abrir la URL, ingresar con ese usuario y contraseña.

Notas:
- En el plan gratuito de Render el servicio se duerme si no se usa; la primera carga tarda más.
- El Excel de stock es grande (~11 MB): el procesamiento puede demorar entre 30 y 60 segundos.
- Si el build falla por la versión de Python, ajustar `PYTHON_VERSION` en `render.yaml`.

## Probar en la computadora

```
pip install -r requirements.txt
APP_USER=admin APP_PASSWORD=clave python app.py     # http://localhost:5000
```

## Pendiente para próximas versiones
- División por embarques (10005-1, 10005-2…).
- Agrupar códigos que se tratan como uno solo (ej. FILETE CC 4/5 y +5 LBS).
