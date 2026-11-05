# Next Session

Pendientes para retomar en la próxima sesión.

## App Judo Hatillo

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

_Anotado: 2026-11-05_
