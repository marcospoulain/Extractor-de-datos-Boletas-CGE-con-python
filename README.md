# Extractor de Boletas CGE

Una aplicación de escritorio que automatiza la extracción de datos desde boletas eléctricas de CGE a archivos Excel, optimizando el proceso de digitalización para empresas y municipios.

## Características

- Interfaz gráfica intuitiva
- Extracción automática de datos clave de las boletas
- Exportación a formato Excel
- Barra de progreso para seguimiento del proceso
- Visualización inmediata de la ubicación del archivo generado

## Datos Extraídos

- Número de Cliente
- RUT
- Número de Factura/Boleta
- Fecha de Emisión
- Total a Pagar
- Fecha de Vencimiento
- Último Pago
- Consumo Total del Mes (kWh)

## Requisitos

- Python 3.6 o superior
- Dependencias listadas en `requirements.txt`

## Instalación

1. Clonar el repositorio:
```bash
git clone [URL del repositorio]
```

2. Instalar dependencias:
```bash
pip install -r requirements.txt
```

## Uso

1. Ejecutar el programa:
```bash
python pdf_extractor.py
```

2. Hacer clic en "Seleccionar PDF" y elegir el archivo de boleta
3. Presionar "Extraer Datos" para iniciar el proceso
4. El archivo Excel se generará en la misma ubicación que el PDF

## Contribución

Las contribuciones son bienvenidas. Por favor, abrir un issue para discutir cambios mayores antes de crear un pull request.

## Licencia

Este proyecto está bajo la Licencia MIT - ver el archivo [LICENSE](LICENSE) para más detalles.
