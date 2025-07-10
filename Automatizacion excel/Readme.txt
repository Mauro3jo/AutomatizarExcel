#

📄 Automatización de Excel Backoffice

---

## 🚀 ¿Qué es este sistema?

Este proyecto automatiza y controla procesos administrativos complejos basados en archivos Excel.
Ideal para equipos de pagos, backoffice financiero y operaciones, elimina tareas manuales, reduce errores y asegura controles automáticos de calidad sobre datos de tarjetas, QR, SAS, crudos, CRM y libros diarios.

El sistema está dividido en **pasos modulares**, cada uno con funciones específicas, y cuenta con controles visuales, reportes y validaciones avanzadas.

---

## 🗂️ Estructura General y Flujo de Trabajo

### 1. Paso 1 — Procesamiento y Control Inicial

* Selección del archivo principal de Excel.
* Procesamiento independiente para cada hoja relevante (Visa, Mastercard, Maestro, etc.).
* Limpieza y normalización de filas según reglas de negocio.
* Suma automática de importes brutos, validación y reporte de errores o filas anómalas.
* Botón de control de tasas y verificación de columnas específicas.
* Avance al siguiente paso solo cuando todo está correcto.

### 2. Paso 2 — Reubicación, Integración y Excepciones

* Selección de un segundo archivo Excel (SAS).
* Reubica operaciones por fecha y controla transferencias entre archivos.
* Permite integrar operaciones desde hojas de excepción, ideal para feriados o movimientos no regulares.
* Muestra estado de procesamiento y reportes detallados.

### 3. Paso 3 — Copia, Fusión y Consolidación

* Carga y control de archivos CRM y Crudo.
* Botones de acción para copiar ALTAS, BAJAS y SAS entre archivos.
* Controla la integridad y finalización de cada proceso antes de avanzar.
* Visibilidad de los archivos procesados y feedback visual.

### 4. Paso 4 — Validaciones Finales y Control Diario

* Carga del archivo Diario.
* Controles automáticos de fechas únicas, suma de brutos, verificación de arancel, IVA y costo transaccional.
* Validación avanzada de IIBB usando parámetros de base de datos.
* Botón exclusivo para validación FUR.
* Feedback visual y reportes centralizados.

### QR — Procesamiento Especial para Datos QR

* Procesamiento y pegado automatizado de datos QR desde distintos archivos.
* Ejecución de macros automatizadas sobre archivos QR/SAS.
* Copia de ALTAS, BAJAS y SAS desde CRM y Crudo.
* Control de rutas y archivos cargados, con validaciones y mensajes de éxito/error.

---

## 🖥️ ¿Cómo instalar y ejecutar?

1. **Clonar el repositorio**

   ```sh
   git clone <TU_REPO>
   ```

2. **Abrir en Visual Studio 2022+**
   Abrí la solución (ej: `Automatizacion_excel.sln`).

3. **Restaurar paquetes y compilar**
   El sistema está basado en **.NET 8.0**.
   Si es la primera vez, Visual Studio te pedirá instalar el SDK si no lo tenés.

4. **Ejecutar**

   * Ejecutá el proyecto principal (WinForms).
   * Seleccioná el flujo que necesitás (Fiserv, QR, etc).
   * Seguí los pasos en pantalla, cargando los archivos que te solicita el sistema.

---

## 🧪 Pruebas y Testing

* El proyecto incluye **pruebas unitarias y de integración** en la carpeta `/Automatizacion.Tests`.
* Para pruebas de integración:

  * Usá archivos de ejemplo (por ejemplo, en `TestFiles/Ejemplo.xlsx`).
  * Los tests cubren desde la carga hasta el procesamiento y control de los archivos.
* Podés crear tus propios archivos Excel de prueba para simular distintos escenarios.
* Las pruebas simulan el flujo completo, verificando que los resultados sean correctos y no se produzcan errores en el proceso.

**Ejemplo de cómo correr los tests:**

* Abrí el proyecto en Visual Studio.
* Usá el **Test Explorer** (`Ctrl+R, A`) para correr todos los tests.
* Revisá los resultados y cobertura.

---

## 📂 Estructura de Carpetas (resumida)

```
Automatizacion_excel/
├── Formularios/                # Formularios WinForms
├── Paso1/                      # Procesos y lógica de Paso 1
├── Paso2/
├── Paso3/
├── Paso4/
├── QR/                         # Lógica específica QR y macros
├── Core/                       # Servicios e interfaces reusables
├── Data/                       # Acceso a base de datos
├── Tests/                      # Pruebas automatizadas
└── TestFiles/                  # Archivos Excel de ejemplo para testing
```

---

## 🏆 Decisiones técnicas y justificación

### Procesamiento individual por tarjeta

Cada hoja/tarjeta se procesa con una clase (“Processor”) específica y no con una clase genérica.
Esto se debe a que **cada banco, tarjeta o proveedor puede cambiar el formato, la ubicación de las columnas o incluso la lógica de negocio de sus archivos Excel en cualquier momento**. Mantener un processor por tarjeta:

* Aísla los cambios y facilita el mantenimiento.
* Permite reglas y validaciones específicas para cada caso, sin riesgo de romper los procesos de otras tarjetas.
* Hace el código más claro y explícito, facilitando la auditoría y la incorporación de nuevos colaboradores.

### Uso de Interop.Excel en lugar de librerías modernas

El sistema utiliza **Microsoft.Office.Interop.Excel** en vez de librerías como EPPlus, ClosedXML o NPOI porque:

* **Permite trabajar con archivos Excel que contienen macros (VBA)**, ejecutar esas macros desde el sistema, y manipular archivos que incluyen lógica o automatización interna.
* Las alternativas modernas no permiten ejecutar ni modificar macros y pueden romper el contenido o la funcionalidad embebida.
* Interop.Excel es la **única opción robusta** cuando el archivo Excel no es “sólo datos”, sino también procesos automáticos o validaciones internas que dependen de macros.

### Arquitectura orientada a pasos y sin MVC/MVVM

La arquitectura se basa en **clases orquestadoras por paso**, en lugar de patrones como MVC/MVVM, porque:

* WinForms y la automatización de Excel requieren acceso directo a recursos de Office y la UI, haciendo innecesario y hasta contraproducente forzar patrones de desacoplamiento estrictos pensados para la web o WPF.
* La separación real se da entre **UI, lógica de negocio (`Core`) y acceso a datos (`Data`)**, manteniendo el sistema flexible y fácil de mantener.
* Si en el futuro se migrara a WPF o web, la lógica del core puede migrarse a servicios o APIs más desacopladas.

### Pruebas automatizadas y cobertura

* Todos los módulos críticos (processors, validadores, servicios) cuentan con **tests automatizados reales**, que usan archivos de ejemplo y verifican resultados de punta a punta.
* Las pruebas aseguran que cualquier cambio en reglas de negocio, formato de archivo o validaciones sea detectado antes de que impacte en producción.

---

## 🏗️ Tecnologías y Dependencias

* **.NET 8.0**
* **WinForms** para interfaz de usuario
* **Microsoft.Office.Interop.Excel** para manipulación de archivos Excel
* **MSTest** para testing unitario e integración

---

## 🔗 Notas para desarrolladores

* El sistema sigue una arquitectura modular y desacoplada: cada Paso es una clase independiente.
* La lógica de negocio está desacoplada de la interfaz gráfica.
* Los eventos (`Paso1Completado`, `Paso2Completado`, etc.) se usan para orquestar el avance entre pasos.
* Para agregar nuevos controles, basta con extender los pasos o crear nuevos módulos en `/Core` y `/Data`.
* Si necesitás migrar a web o móvil, la lógica central se puede reusar en backend.

---

## 🤝 Contribuciones y mejoras

1. **Abrí un Issue** con tu sugerencia o bug.
2. **Forkeá** y trabajá en una rama propia.
3. Hacé un Pull Request describiendo los cambios.
4. Mantené la estructura modular y buenas prácticas.

---

## 👤 Autor

* Desarrollador principal: Trejo Mauro
* Empresa: \[Zoco Servicios de Pago / ZOCO]

---

## ❓ Preguntas frecuentes

**¿Puedo correr esto en la web?**
No. Está pensado para escritorio Windows (WinForms).

**¿Qué archivos Excel usar?**
Los que usa tu operación diaria, o los de ejemplo para pruebas.

**¿Puedo agregar nuevas tarjetas/procesos?**
Sí. Duplicá la lógica en Paso1 o creá nuevos módulos.

**¿Se puede testear todo automáticamente?**
Sí, pero debés tener archivos Excel de ejemplo y configurar rutas relativas en los tests para evitar errores de permisos o acceso.

---

## 📝 Licencia

Licencia interna para uso de Zoco Servicios de Pago / ZOCO.

