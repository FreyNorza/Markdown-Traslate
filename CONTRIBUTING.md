# Guía de Contribución 🤝

¡Gracias por tu interés en contribuir a **Markdown Converter**!

## ¿Cómo puedes colaborar?

1. **Reportar errores (Bugs):**
   - Antes de abrir un Issue, revisa que no haya uno existente sobre el mismo tema.
   - Describe con claridad el formato de entrada, el comportamiento esperado y el mensaje de error o aviso que se produjo.

2. **Proponer mejoras o nuevos formatos:**
   - Si crees que un formato adicional o una opción de estilo en Word aportaría valor, crea un Issue explicando el caso de uso.

3. **Enviar código (Pull Requests):**
   - Asegúrate de seguir la estructura limpia del proyecto:
     - Lógica de conversión pura en `converters/` (sin código UI).
     - Componentes visuales en `ui/`.
     - Pruebas sintácticas antes de enviar tu PR.
   - Mantén los commits descriptivos en español o inglés siguiendo buenas prácticas (ej. `feat: ...`, `fix: ...`, `docs: ...`).

---

## Entorno de desarrollo local

```bat
# Activar entorno
.venv\Scripts\activate

# Instalar dependencias
pip install -r requirements.txt

# Ejecutar aplicación
python main.py
```

