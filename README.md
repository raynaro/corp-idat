# Sistema IDAT v6 - Panel de Seguridad mejorado

## Ejecutar en Windows

```cmd
cd C:\Ruta\donde\descomprimiste\sistema_idat_v3
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
python app.py
```

Abrir en navegador:

```txt
http://localhost:5000
```

## Usuarios de prueba

- Administrador: admin@idat.edu.pe / admin123
- SAE: sae@idat.edu.pe / sae123
- Ventas: ventas@idat.edu.pe / ventas123
- Seguridad: seguridad@idat.edu.pe / seguridad123

## Mejoras v6

- Panel especial para Seguridad en `/seguridad`.
- Seguridad puede registrar alumnos de forma rápida.
- Seguridad puede registrar docentes por DNI.
- Seguridad puede registrar visitas directas para SAE o Ventas con botones separados.
- SAE solo ve visitas destinadas a SAE.
- Ventas solo ve visitas destinadas a Ventas.
- Hora configurada con zona horaria Lima/Perú.


## Versión 7 - SAE y Ventas
- Los apartados de SAE y Ventas muestran las visitas enumeradas por orden de llegada.
- Cada área ve solamente sus propias visitas.
- Se agregó botón "Activar sonido" para avisar cuando Seguridad registre una nueva visita para esa área.
- Por seguridad del navegador, el sonido se activa después de presionar el botón una vez.
