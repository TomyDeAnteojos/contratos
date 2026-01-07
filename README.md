# Generador de Contratos (UNPilar) — HTML/CSS/JS (Word + Vista previa)

- Genera **Word (.docx)** a partir de una plantilla con tags tipo: `&APELLIDO&`, `&NOMBRE&`, `&DNI&`, etc.
- Muestra una **vista previa** en pantalla (texto del contrato + reemplazos).  
  *El Word final mantiene el formato real de la plantilla.*

## Tags soportados
`&APELLIDO&`, `&NOMBRE&`, `&DNI&`, `&CALLE&`, `&NUMERO&`, `&LOCALIDAD&`, `&CUIT&`, `&DOMICILIO&`, `&TELEFONO&`, `&EMAIL&`

> Nota: el anexo del modelo trae `&NOMBRE &` (con un espacio antes del `&` final). Ya está contemplado.

## Cómo abrir

### Opción recomendada: servidor local
Algunos navegadores bloquean `fetch("template.docx")` si abrís el HTML con doble click.

**Windows:**
```bash
py -m http.server 8000
```

**Linux/Mac:**
```bash
python3 -m http.server 8000
```

Luego abrí:
- http://localhost:8000

### Alternativa: abrir `index.html` con doble click
Funciona, pero si “Usar plantilla incluida” falla, subí el `.docx` manualmente con el selector de archivo.
