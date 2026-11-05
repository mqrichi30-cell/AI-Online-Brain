# Next Session

Pendientes para retomar en la próxima sesión.

## App Judo Hatillo

### 🐞 Bugs
- [ ] **"Ver contraseña" no funciona**: el botón/ícono para mostrar la contraseña no la muestra. Revisar en login y registro (y en el link de invitación si tiene campo de contraseña). Probar en celular y en escritorio.

### Cuota mensual del sensei
- [ ] **Verificar primero** si esto ya existe en la app (no estamos seguros de si ya es posible).
- [ ] Al **crear la cuenta**, el sensei debe indicar **cuánto cobra por mes** (campo obligatorio en el registro).
- [ ] Ese monto debe ser **editable desde Configuraciones**.
- [ ] El monto de la mensualidad también lo puede **editar el admin del club**.
- [ ] **Día de cobro sin fecha fija:** depende del **último pago de cada atleta** (ej. si pagó el 12, el próximo cobro vence el 12 del mes siguiente).

### Contrato por atleta y becas
- [ ] Cada atleta tiene un **monto de contrato fijo**, guardado en su contrato. Si la mensualidad del club cambia, **lo que paga el atleta no cambia**.
- [ ] La **beca debe mostrar el monto del contrato** (la mensualidad base que se usó para calcularla).
  - Ejemplo: mensualidad 10.000 + beca 50% → paga **5.000**. Si la mensualidad sube a 25.000, ese atleta **sigue pagando 5.000**.
- [ ] **Lo único que cambia el monto de un atleta** es que se le asigne una beca o que la pierda. Las dos cosas son una **actualización de contrato**.
  - En una actualización de contrato, el nuevo monto se calcula con la **mensualidad vigente en ese momento**.
  - Ejemplo: mensualidad actual 25.000, el atleta pierde la beca → nuevo contrato de **25.000**. Si se le asigna una beca del 50% → **12.500**.
- [ ] **Actualizar la lógica de cada cobro mensual**: cobrar según el **monto del contrato de cada atleta**, no según la mensualidad actual del club.
- [ ] Guardar el historial de contratos (fecha, mensualidad base, % de beca, monto final).

### Invitar atletas por correo
- [ ] El sensei puede **agregar atletas con su correo**.
- [ ] Al agregarlo, se le **envía un link** al atleta para que **cree su cuenta**.

### Foto del atleta
- [ ] Pedirle al atleta que suba una **foto tipo pasaporte** (de frente, fondo claro, solo rostro y hombros).
- [ ] Pedirla al crear la cuenta (desde el link de invitación) y permitir cambiarla después en su perfil.

### Rutinas escritas como nota (con la API de Claude)
- [ ] **Cualquier entrenador** (no solo el sensei) puede crear o editar rutinas.
- [ ] Hay dos formas de **crear o editar** una rutina; la IA es **opcional**:
  1. **Manual (sin IA):** el entrenador **busca los ejercicios** y arma la rutina él mismo, o edita desde **"Editar ejercicio"** (lo que ya existe).
  2. **Con una nota (IA):** el entrenador escribe libremente y Claude la interpreta.
- [ ] **Antes de escribir la nota**, el entrenador debe **elegir una opción** (obligatorio):
  - **Construir rutina nueva:** Claude arma la rutina completa a partir de la nota.
  - **Editar parcialmente:** se elige una rutina existente y Claude aplica **solo los cambios** de la nota (agregar, quitar o modificar ejercicios), sin tocar el resto.
- [ ] Claude devuelve la rutina en el formato de la app (ejercicios, series, repeticiones, tiempos, etc.).
- [ ] El entrenador **revisa y corrige** el resultado antes de guardar. En la edición parcial, mostrar **qué cambió** respecto a la rutina original.
- [ ] Después sigue el flujo que ya existe: elegir **asignar a todos / a algunos**, **fecha** y lo demás que ya pide la app.
- [ ] Técnico:
  - La **API key de Claude va solo en el servidor** (variable de entorno, ej. `ANTHROPIC_API_KEY` en Vercel), **nunca en el frontend** ni en el repo.
  - Usar **salida estructurada (JSON)** que coincida con el modelo de datos de rutina de la app.
  - Manejar errores (nota vacía, texto que no se entiende) con un mensaje claro y la opción de crear la rutina a mano.

_Anotado: 2026-11-05_
