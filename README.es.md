# VSKeyExtractor

VSKeyExtractor es una pequeña herramienta de línea de comandos que intenta extraer la clave de licencia utilizada para activar una instalación local de Visual Studio.

Recorre entradas conocidas de productos de Visual Studio, lee los datos de licencia cifrados del registro de Windows, los descifra en el contexto del usuario actual y busca una clave con formato `AAAAA-BBBBB-CCCCC-DDDDD-EEEEE`.

## Características

- Lee datos de licencia de Visual Studio desde el registro de Windows
- Admite varias generaciones de productos de Visual Studio
- Intenta descifrar blobs de licencia protegidos con el perfil del usuario actual
- Muestra en la consola cualquier valor que parezca una clave

## Requisitos

- Windows
- Una instalación de Visual Studio con datos de licencia locales disponibles
- El runtime de .NET necesario para el proyecto que quieras ejecutar

## Cómo funciona

La herramienta comprueba valores del registro en las rutas de licencia conocidas de Visual Studio y después intenta descifrar los datos almacenados. Si encuentra una clave coincidente en el texto descifrado, la escribe en la salida estándar.

Si no encuentra ninguna clave, simplemente continúa con la siguiente entrada de producto conocida.

## Ejecutar la aplicación

Puedes ejecutar el proyecto directamente con `dotnet run`.

### Destino .NET Framework / Windows

```powershell
dotnet run --project .\VSKeyExtractor.csproj
```

### Proyecto estilo SDK

```powershell
dotnet run --project .\VSKeyExtractor_net.csproj
```

### Ejecutar un framework de destino específico

Si usas el proyecto estilo SDK y quieres apuntar a un framework concreto, puedes usar:

```powershell
dotnet run --project .\VSKeyExtractor_net.csproj -f net8.0-windows
```

Otros frameworks de destino admitidos pueden incluir `net9.0-windows` o `net10.0-windows`, según el SDK de .NET instalado.

## Ejemplo de salida

```text
Found key for Visual Studio 2022 Professional: ABCDE-FGHIJ-KLMNO-PQRST-UVWXY
```

## Notas

- Ejecuta la herramienta en el equipo donde se activó Visual Studio.
- La herramienta solo puede leer datos disponibles para el perfil de usuario actual.
- Algunas entradas pueden no estar disponibles, estar dañadas o protegidas; esto es esperable.

## Traducciones

- [English](README.md)
- [Deutsch](README.de.md)
- [Français](README.fr.md)
