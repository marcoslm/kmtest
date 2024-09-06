# KM-Test

KM-Test es una aplicación sencilla y eficaz para probar el funcionamiento de teclados y ratones, mostrando en tiempo real las teclas presionadas, el estado de los LEDs del teclado (Bloq Num, Bloq Mayús y Bloq Despl), y los eventos del ratón, junto con las coordenadas de movimiento. Es muy útil para por ejemplo diagnosticar membranas de teclado defectuosas.

![kmtest](https://github.com/user-attachments/assets/24820d67-b332-420b-abb3-2dc1704985c4)

## Descripción

La aplicación permite capturar todos los eventos generados por el teclado y el ratón. Las pulsaciones del teclado se muestran visualmente iluminando en rojo las teclas correspondientes en el teclado ilustrado en la interfaz de la aplicación. También se visualiza el estado de los LEDs del teclado (Bloqueo Numérico, Mayúsculas y Desplazamiento). Además, registra las pulsaciones de los botones del ratón, el desplazamiento de las ruedas y las coordenadas de movimiento del cursor en la pantalla.

Este proyecto lo desarrollé hace bastantes años en VB6 y recientemente lo porté a **TwinBASIC**, que es un lenguaje 100% compatible con VB6/VBA con mejoras. Puedes conocer más sobre este lenguaje y su desarrollo en su [repositorio oficial de GitHub](https://github.com/twinbasic/twinbasic).

## Características

- Captura en tiempo real de los eventos de teclado y ratón.
- Visualización de las teclas presionadas con iluminación en rojo sobre el teclado gráfico.
- Registro del estado de los LEDs del teclado (Bloq Num, Bloq Mayús, Bloq Despl).
- Registro de las pulsaciones de botones del ratón (derecho, izquierdo, central) y la rueda de desplazamiento.
- Muestra las coordenadas exactas del ratón (ejes X e Y).

## Requisitos del Sistema

- Sistema operativo: Windows 7 o superior.
- No requiere instalación, simplemente descargar y ejecutar.

## Uso

1. Descarga el ejecutable desde [aquí](https://github.com/marcoslm/kmtest/raw/main/Build/kmtest_win32.exe).
2. Ejecuta el archivo directamente, no se requiere instalación.
3. Haz clic en "Comenzar" para empezar a capturar eventos del teclado y ratón.
4. A partir de ese momento, cada pulsación de tecla o clic del ratón será registrado en la ventana de eventos.
5. Observa cómo se iluminan en rojo las teclas que presionas en el teclado visual de la interfaz y cómo se reflejan los cambios en los LEDs del teclado (Num, Mayús, Despl).
6. Los eventos del ratón, como clics y movimientos, también se mostrarán en tiempo real junto con las coordenadas del cursor.

## Licencia

Este proyecto está bajo la [Licencia Creative Commons Attribution-NonCommercial 4.0 International (CC BY-NC 4.0)]([LICENSE](https://github.com/marcoslm/kmtest/raw/main/LICENSE)).

