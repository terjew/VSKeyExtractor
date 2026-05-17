# VSKeyExtractor

VSKeyExtractor is a small command-line tool that attempts to extract the license key used to activate a local Visual Studio installation.

It scans known Visual Studio product entries, reads the encrypted license data from the Windows registry, decrypts it with the current user context, and looks for a key formatted like `AAAAA-BBBBB-CCCCC-DDDDD-EEEEE`.

## Features

- Reads Visual Studio license data from the Windows registry
- Supports multiple Visual Studio product generations
- Tries to decrypt protected license blobs with the current user profile
- Prints any key-shaped values it finds to the console

## Requirements

- Windows
- Visual Studio installation with local license data present
- The required .NET runtime for the project you want to run

## How it works

The tool checks registry values under the Visual Studio licensing locations, then tries to decrypt the stored data. If a matching key is found in the decrypted text, it is written to standard output.

If no key is found, the tool simply continues with the next known product entry.

## Run the application

You can run the project directly with `dotnet run`.

### .NET Framework / Windows target

```powershell
dotnet run --project .\VSKeyExtractor.csproj
```

### SDK-style project

```powershell
dotnet run --project .\VSKeyExtractor_net.csproj
```

### Run a specific target framework

If you are using the SDK-style project and want to target a specific framework, you can use:

```powershell
dotnet run --project .\VSKeyExtractor_net.csproj -f net8.0-windows
```

Other supported target frameworks may include `net9.0-windows` or `net10.0-windows`, depending on the installed .NET SDK.

## Example output

```text
Found key for Visual Studio 2022 Professional: ABCDE-FGHIJ-KLMNO-PQRST-UVWXY
```

## Notes

- Run the tool on the machine where Visual Studio was activated.
- The tool can only read data that is available to the current user profile.
- Some entries may be unavailable, malformed, or protected, which is expected.

## Translations

- [Deutsch](README.de.md)
- [Français](README.fr.md)
- [Español](README.es.md)
